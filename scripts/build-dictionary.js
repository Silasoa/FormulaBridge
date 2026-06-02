import fs from 'fs';
import path from 'path';
import os from 'os';
import { execSync } from 'child_process';

import { fileURLToPath } from 'url';

// Define relative paths for security (no hardcoded absolute paths)
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const workspaceDir = path.resolve(__dirname, '..');
const existingDictPath = path.join(workspaceDir, 'src/data/dictionary.json');
const htmlSourcePath = path.join(__dirname, 'source.html');
const outputDictPath = path.join(workspaceDir, 'src/data/dictionary.json');

// LCID mappings to our target languages
const lcidMap = {
  1029: "cs",
  1030: "da",
  1031: "de",
  1035: "fi",
  1036: "fr",
  1038: "hu",
  1040: "it",
  1043: "nl",
  1044: "nb",
  1045: "pl",
  1046: "pt-BR",
  1049: "ru",
  1053: "sv",
  1055: "tr",
  3082: "es"
};

async function main() {
  console.log('Starting dictionary builder...');
  console.log(`Looking for HTML source at: ${htmlSourcePath}`);

  // 1. Hardcoded Error Codes and Booleans
  const specialEntries = {
    "TRUE": { "de": "WAHR", "fr": "VRAI", "es": "VERDADERO", "it": "VERO", "pt-BR": "VERDADEIRO", "nl": "WAAR", "pl": "PRAWDA", "cs": "PRAVDA", "ru": "ИСТИНА", "tr": "DOĞRU", "sv": "SANT", "nb": "SANN", "da": "SAND", "fi": "TOSI", "hu": "IGAZ", "sk": "PRAVDA" },
    "FALSE": { "de": "FALSCH", "fr": "FAUX", "es": "FALSO", "it": "FALSO", "pt-BR": "FALSO", "nl": "ONWAAR", "pl": "FAŁSZ", "cs": "NEPRAVDA", "ru": "ЛОЖЬ", "tr": "YANLIŞ", "sv": "FALSKT", "nb": "USANN", "da": "FALSK", "fi": "EPÄTOSI", "hu": "HAMIS", "sk": "NEPRAVDA" },
    "#DIV/0!": { "de": "#DIV/0!", "fr": "#DIV/0!", "es": "#¡DIV/0!", "it": "#DIV/0!", "pt-BR": "#DIV/0!", "nl": "#DEEL/0!", "pl": "#DZIEL/0!", "cs": "#DĚLENÍ_NULOU!", "ru": "#ДЕЛ/0!", "tr": "#SAYI/0!", "sv": "#DIVISION/0!", "nb": "#DIV/0!", "da": "#DIVISION/0!", "fi": "#JAKO/0!", "hu": "#ZÉRÓOSZTÓ!", "sk": "#DELENIE_NULOU!" },
    "#N/A": { "de": "#NV", "fr": "#N/A", "es": "#N/A", "it": "#N/D", "pt-BR": "#N/D", "nl": "#N/B", "pl": "#N/D!", "cs": "#NENÍ_K_DISPOZICI", "ru": "#Н/Д", "tr": "#YOK", "sv": "#SAKNAS!", "nb": "#I/T", "da": "#I/T", "fi": "#PUUTTUU!", "hu": "#HIÁNYZIK", "sk": "#NEDOSTUPNÉ" },
    "#NAME?": { "de": "#NAME?", "fr": "#NOM?", "es": "#¿NOMBRE?", "it": "#NOME?", "pt-BR": "#NOME?", "nl": "#NAAM?", "pl": "#NAZWA?", "cs": "#NÁZEV?", "ru": "#ИМЯ?", "tr": "#AD?", "sv": "#NAMN?", "nb": "#NAVN?", "da": "#NAVN?", "fi": "#NIMI?", "hu": "#NÉV?", "sk": "#NÁZOV?" },
    "#NULL!": { "de": "#NULL!", "fr": "#NUL!", "es": "#¡NULO!", "it": "#NULLO!", "pt-BR": "#NULO!", "nl": "#LEEG!", "pl": "#PUSTE!", "cs": "#HODNOTA!", "ru": "#ПУСТО!", "tr": "#BOŞ!", "sv": "#NULL!", "nb": "#NULL!", "da": "#NULL!", "fi": "#PUUTTUU!", "hu": "#NULLA!", "sk": "#NULL!" },
    "#NUM!": { "de": "#ZAHL!", "fr": "#NOMBRE!", "es": "#¡NUM!", "it": "#NUM!", "pt-BR": "#NÚM!", "nl": "#GETAL!", "pl": "#LICZBA!", "cs": "#ČÍSLO!", "ru": "#ЧИСЛО!", "tr": "#SAYI!", "sv": "#OGILTIGT!", "nb": "#NUM!", "da": "#NUM!", "fi": "#LUKU!", "hu": "#SZÁM!", "sk": "#ČÍSLO!" },
    "#REF!": { "de": "#BEZUG!", "fr": "#REF!", "es": "#¡REF!", "it": "#RIF!", "pt-BR": "#REF!", "nl": "#VERW!", "pl": "#ADR!", "cs": "#ODKAZ!", "ru": "#ССЫЛКА!", "tr": "#BAŞV!", "sv": "#REFERENS!", "nb": "#REF!", "da": "#REF!", "fi": "#VIITTAUS!", "hu": "#HIV!", "sk": "#ODKAZ!" },
    "#VALUE!": { "de": "#WERT!", "fr": "#VALEUR!", "es": "#¡VALOR!", "it": "#VALORE!", "pt-BR": "#VALOR!", "nl": "#WAARDE!", "pl": "#ARG!", "cs": "#HODNOTA!", "ru": "#ЗНАЧ!", "tr": "#DEĞER!", "sv": "#VÄRDEFEL!", "nb": "#VERDI!", "da": "#VÆRDI!", "fi": "#ARVO!", "hu": "#ÉRTÉK!", "sk": "#HODNOTA!" },
    "#GETTING_DATA": { "de": "#DATEN_ABRUFEN", "fr": "#OBTENTION_DONNEES", "es": "#OBTENIENDO_DATOS", "it": "#RECUPERO_DATI", "pt-BR": "#OBTENDO_DADOS", "nl": "#GEGEVENS_OPHALEN", "pl": "#POBIERANIE_DANYCH", "cs": "#ZÍSKÁVÁNÍ_DAT", "ru": "#ПОЛУЧЕНИЕ_ДАННЫХ", "tr": "#VERİ_ALINIYOR", "sv": "#HÄMTAR_DATA", "nb": "#HENTER_DATA", "da": "#HENTER_DATA", "fi": "#NOUDETAAN_TIETOJA", "hu": "#ADATOK_BEKÉRÉSE", "sk": "#ZÍSKAVAJÚ_SA_ÚDAJE" }
  };

  let existingLocales = {};
  if (fs.existsSync(existingDictPath)) {
    try {
      const existingData = JSON.parse(fs.readFileSync(existingDictPath, 'utf8'));
      existingLocales = existingData.locales || {};
    } catch (e) {
      console.warn('Could not read existing locales:', e.message);
    }
  }

  // 2. Read HTML source
  if (!fs.existsSync(htmlSourcePath)) {
    console.error(`HTML source file not found at ${htmlSourcePath}`);
    process.exit(1);
  }
  const htmlContent = fs.readFileSync(htmlSourcePath, 'utf8');
  console.log(`Read HTML source file (${htmlContent.length} chars)`);

  // 3. Parse HTML using regular expressions
  const rowRegex = /<tr>([\s\S]*?)<\/tr>/g;
  let match;
  const functions = {};

  const tdRegex = /<td[^>]+id="([0-9]+)R([0-9]+)C([0-9]+)"[^>]*>([\s\S]*?)<\/td>/g;

  let rowCount = 0;
  let functionCount = 0;

  while ((match = rowRegex.exec(htmlContent)) !== null) {
    const rowHtml = match[1];
    
    const cells = {};
    let tdMatch;
    tdRegex.lastIndex = 0; // Reset regex
    
    while ((tdMatch = tdRegex.exec(rowHtml)) !== null) {
      const lcid = parseInt(tdMatch[1], 10);
      let cellText = tdMatch[4].trim();

      // Clean up cell text: strip tags like <a ...>...</a>
      cellText = cellText.replace(/<[^>]+>/g, '').trim();
      
      cells[lcid] = cellText;
    }

    if (cells[1033]) {
      const enName = cells[1033].toUpperCase();
      
      if (!enName || enName === 'A' || enName === 'B' || enName === 'C' || enName === 'PRODUCT' || enName === 'SOURCE LANGUAGE') {
        continue;
      }

      const translations = {};
      
      for (const [lcid, langCode] of Object.entries(lcidMap)) {
        if (cells[lcid]) {
          translations[langCode] = cells[lcid].trim();
        } else {
          translations[langCode] = enName;
        }
      }

      // If we don't have existing sk data for the HTML parsed ones, just use EN
      translations['sk'] = enName;

      functions[enName] = translations;
      functionCount++;
    }
    rowCount++;
  }

  console.log(`Parsed ${rowCount} rows, found ${functionCount} valid function entries from HTML.`);

  // 4. Merge errors, booleans
  let restoredCount = 0;
  for (const [key, trans] of Object.entries(specialEntries)) {
    functions[key] = trans;
    restoredCount++;
  }
  console.log(`Injected ${restoredCount} special keys (errors, booleans) into the dictionary.`);

  // Sort keys alphabetically
  const sortedFunctions = {};
  Object.keys(functions).sort().forEach(key => {
    sortedFunctions[key] = functions[key];
  });

  const finalDict = {
    meta: {
      version: "1.0.0",
      totalFunctions: Object.keys(sortedFunctions).length,
      languages: ["de", "fr", "es", "it", "pt-BR", "nl", "pl", "cs", "ru", "tr", "sv", "nb", "da", "fi", "hu", "sk"]
    },
    locales: existingLocales,
    functions: sortedFunctions
  };

  fs.writeFileSync(outputDictPath, JSON.stringify(finalDict, null, 2), 'utf8');
  console.log(`Successfully wrote expanded dictionary to ${outputDictPath} with total ${Object.keys(sortedFunctions).length} entries.`);
}

main().catch(err => {
  console.error('Error building dictionary:', err);
  process.exit(1);
});
