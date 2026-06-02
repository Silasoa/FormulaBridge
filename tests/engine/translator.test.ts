// ============================================================
// FormulaBridge – Unit Tests: Translator
// ============================================================

import { describe, it, expect } from "vitest";
import { translateFormula } from "../../src/engine/translator.js";
import { Dictionary } from "../../src/engine/types.js";
import dictionaryData from "../../src/data/dictionary.json";

const dictionary = dictionaryData as unknown as Dictionary;

describe("Translation Engine", () => {
  // ── DE → EN ──
  it("should translate simple German formula to English", () => {
    const result = translateFormula("=SUMME(A1:A10)", dictionary, "en", "de");
    expect(result.translated).toBe("=SUM(A1:A10)");
    expect(result.wasTranslated).toBe(true);
    expect(result.changes).toHaveLength(1);
    expect(result.changes[0]).toMatchObject({ original: "SUMME", translated: "SUM" });
  });

  it("should translate nested German formula to English", () => {
    const result = translateFormula(
      "=WENN(SUMME(A1:A10)>0;MITTELWERT(B1:B10);0)",
      dictionary, "en", "de"
    );
    expect(result.translated).toBe("=IF(SUM(A1:A10)>0,AVERAGE(B1:B10),0)");
    expect(result.wasTranslated).toBe(true);
    expect(result.changes).toHaveLength(3);
  });

  it("should protect strings from translation", () => {
    const result = translateFormula(
      '=VERKETTEN("SUMME";A1)',
      dictionary, "en", "de"
    );
    expect(result.translated).toBe('=CONCATENATE("SUMME",A1)');
    // "SUMME" inside quotes should NOT be translated
    expect(result.translated).toContain('"SUMME"');
  });

  it("should handle mixed-language formulas", () => {
    const result = translateFormula(
      "=SUMME(AVERAGE(A1:A10))",
      dictionary, "en", "de"
    );
    // SUMME should be translated, AVERAGE is already English
    expect(result.translated).toBe("=SUM(AVERAGE(A1:A10))");
  });

  // ── FR → EN ──
  it("should translate French formula to English", () => {
    const result = translateFormula("=SOMME(A1:A10)", dictionary, "en", "fr");
    expect(result.translated).toBe("=SUM(A1:A10)");
  });

  it("should translate French SI to English IF with separator", () => {
    const result = translateFormula("=SI(A1>0;A1;0)", dictionary, "en", "fr");
    expect(result.translated).toBe("=IF(A1>0,A1,0)");
  });

  // ── Separator conversion ──
  it("should convert ; to , when translating to English", () => {
    const result = translateFormula("=WENN(A1>0;A1;0)", dictionary, "en", "de");
    expect(result.translated).toBe("=IF(A1>0,A1,0)");
  });

  it("should convert , to ; when translating to German", () => {
    const result = translateFormula("=IF(A1>0,A1,0)", dictionary, "de", "en");
    expect(result.translated).toBe("=WENN(A1>0;A1;0)");
  });

  // ── Unknown functions ──
  it("should leave unknown functions unchanged", () => {
    const result = translateFormula("=MEINEFUNKTION(A1)", dictionary, "en", "de");
    expect(result.translated).toBe("=MEINEFUNKTION(A1)");
  });

  // ── Cell references ──
  it("should not modify cell references", () => {
    const result = translateFormula("=SUMME(Sheet1!$A$1:$B$10)", dictionary, "en", "de");
    expect(result.translated).toBe("=SUM(Sheet1!$A$1:$B$10)");
  });

  // ── No translation needed ──
  it("should return wasTranslated=false when source equals target", () => {
    const result = translateFormula("=SUM(A1:A10)", dictionary, "en", "en");
    expect(result.wasTranslated).toBe(false);
    expect(result.translated).toBe("=SUM(A1:A10)");
  });

  it("should not translate non-formula strings", () => {
    const result = translateFormula("Hello World", dictionary, "en", "de");
    expect(result.wasTranslated).toBe(false);
    expect(result.translated).toBe("Hello World");
  });

  // ── Auto-detect ──
  it("should auto-detect German and translate", () => {
    const result = translateFormula(
      "=WENN(ODER(A1>0;A2>0);SUMME(A1:A2);0)",
      dictionary, "en"
    );
    expect(result.sourceLanguage).toBe("de");
    expect(result.translated).toBe("=IF(OR(A1>0,A2>0),SUM(A1:A2),0)");
    expect(result.confidence).toBeGreaterThan(0.5);
  });

  it("should auto-detect French and translate", () => {
    const result = translateFormula("=SOMME(A1:A10)", dictionary, "en");
    expect(result.translated).toBe("=SUM(A1:A10)");
  });

  // ── EN → DE (reverse) ──
  it("should translate English to German", () => {
    const result = translateFormula("=VLOOKUP(A1;B1:C10;2;FALSE)", dictionary, "de", "en");
    expect(result.translated).toBe("=SVERWEIS(A1;B1:C10;2;FALSCH)");
  });

  // ── Complex formulas ──
  it("should handle complex nested formula", () => {
    const result = translateFormula(
      "=WENN(UND(A1>0;ODER(B1>0;C1>0));SUMME(A1:C1);0)",
      dictionary, "en", "de"
    );
    expect(result.translated).toBe("=IF(AND(A1>0,OR(B1>0,C1>0)),SUM(A1:C1),0)");
  });

  // ── Error Literals ──
  it("should translate error literals between English and German", () => {
    // English to German
    const toDe = translateFormula("=IFERROR(A1;#N/A)", dictionary, "de", "en");
    expect(toDe.translated).toBe("=WENNFEHLER(A1;#NV)");
    expect(toDe.changes).toContainEqual(expect.objectContaining({ original: "IFERROR", translated: "WENNFEHLER" }));
    expect(toDe.changes).toContainEqual(expect.objectContaining({ original: "#N/A", translated: "#NV" }));

    // German to English
    const toEn = translateFormula("=IFERROR(A1;#NV)", dictionary, "en", "de");
    expect(toEn.translated).toBe("=IFERROR(A1,#N/A)");
  });

  it("should translate other error literals correctly", () => {
    const result = translateFormula("=IF(A1=1;#REF!;#VALUE!)", dictionary, "de", "en");
    expect(result.translated).toBe("=WENN(A1=1;#BEZUG!;#WERT!)");
  });

  // ── Multi-language UTF-8 Coverage ──
  it("should correctly translate between English and all supported languages (checking UTF-8)", () => {
    const testCases: Record<string, { en: string; localized: string }> = {
      "de": { en: "=AVERAGE(A1:A10)", localized: "=MITTELWERT(A1:A10)" },
      "fr": { en: "=AVERAGE(A1:A10)", localized: "=MOYENNE(A1:A10)" },
      "es": { en: "=AVERAGE(A1:A10)", localized: "=PROMEDIO(A1:A10)" },
      "it": { en: "=AVERAGE(A1:A10)", localized: "=MEDIA(A1:A10)" },
      "pt-BR": { en: "=AVERAGE(A1:A10)", localized: "=MÉDIA(A1:A10)" }, // UTF-8 Check (É)
      "nl": { en: "=AVERAGE(A1:A10)", localized: "=GEMIDDELDE(A1:A10)" },
      "pl": { en: "=AVERAGE(A1:A10)", localized: "=ŚREDNIA(A1:A10)" },   // UTF-8 Check (Ś)
      "cs": { en: "=AVERAGE(A1:A10)", localized: "=PRŮMĚR(A1:A10)" },    // UTF-8 Check (Ů, Ě, Ř)
      "ru": { en: "=AVERAGE(A1:A10)", localized: "=СРЗНАЧ(A1:A10)" },    // UTF-8 Check (Cyrillic)
      "tr": { en: "=AVERAGE(A1:A10)", localized: "=ORTALAMA(A1:A10)" },
      "sv": { en: "=AVERAGE(A1:A10)", localized: "=MEDEL(A1:A10)" },
      "nb": { en: "=AVERAGE(A1:A10)", localized: "=GJENNOMSNITT(A1:A10)" },
      "da": { en: "=AVERAGE(A1:A10)", localized: "=MIDDEL(A1:A10)" },
      "fi": { en: "=AVERAGE(A1:A10)", localized: "=KESKIARVO(A1:A10)" },
      "hu": { en: "=AVERAGE(A1:A10)", localized: "=ÁTLAG(A1:A10)" },      // UTF-8 Check (Á)
      "sk": { en: "=AVERAGE(A1:A10)", localized: "=PRIEMER(A1:A10)" }
    };

    for (const [lang, formulaPair] of Object.entries(testCases)) {
      // English to Localized (Note: English uses comma, localized uses semicolon)
      const toLocalizedFormula = formulaPair.en;
      const expectedLocalizedFormula = formulaPair.localized.replace(",", ";");
      const resultTo = translateFormula(toLocalizedFormula, dictionary, lang as any, "en");
      expect(resultTo.translated, `Failed translating EN -> ${lang}`).toBe(expectedLocalizedFormula);

      // Localized to English
      const resultFrom = translateFormula(formulaPair.localized, dictionary, "en", lang as any);
      expect(resultFrom.translated, `Failed translating ${lang} -> EN`).toBe(formulaPair.en);
    }
  });
});
