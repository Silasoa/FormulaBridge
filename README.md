# 🌉 FormulaBridge

> *"Your formulas. Your language. Every Excel."*

**FormulaBridge** is a modern, real-time Excel Add-in that translates Excel formulas between 18 languages automatically. It bridges the gap between different Excel localization settings by translating function names (like `SUM` to `SUMME`) and separators (like `,` to `;`) dynamically as you type.

**Important Note:** This project was built via **Vibe Coding with Antigravity** (Google DeepMind) as a learning project for modern Agentic AI development. Use at your own risk! 🚀

---

## ✨ Features

- **Real-Time Translation**: Attach to the `onChanged` event to instantly translate formulas upon pressing Enter.
- **18 Languages Supported**: English, German, French, Spanish, Italian, Portuguese (BR), Dutch, Polish, Czech, Russian, Turkish, Swedish, Norwegian Bokmål, Danish, Finnish, Hungarian, and Slovak.
- **Intelligent Parsing**: A custom tokenizer ensures string literals (e.g. `="SUM"`) and sheet names are never accidentally translated.
- **Copy-Paste Translator**: Bypass Excel's strict syntax validation (which blocks pasting formulas with foreign separators) by pasting them directly into the Add-in Taskpane.
- **Formula Reference**: A built-in, searchable dictionary mapping all 494 functions between English and your local target language.
- **Zero UI Framework**: Built entirely on Vanilla TypeScript and Fluent UI Web Components for maximum performance and zero bloat (`0KB` framework overhead).

---

## 🛠️ Tech Stack

FormulaBridge is built to be extremely lightweight and fast. It deliberately avoids heavy UI frameworks like React or Angular, because an Excel Taskpane should open instantly and consume minimal memory.

- **UI Framework**: Vanilla TypeScript + DOM API
- **Design System**: [Fluent UI Web Components](https://learn.microsoft.com/en-us/fluent-ui/web-components/)
- **Bundler**: [Vite 6](https://vitejs.dev/)
- **Office API**: `Office.js` (Excel JavaScript API)
- **Manifest**: Unified JSON Manifest (`manifest.json`)
- **Testing**: Vitest

---

## 📦 Installation & Sideloading

To run FormulaBridge locally in your Excel Desktop or Excel Online environment:

### 1. Install Dependencies
```bash
npm install
```

### 2. Generate Development Certificates
Office Add-ins require HTTPS. Generate the necessary self-signed certificates:
```bash
npx office-addin-dev-certs install
```

### 3. Start the Development Server
```bash
npm run dev
```
The server will start at `https://localhost:3000`.

### 4. Sideload into Excel
- **Windows Desktop**: 
  1. Open a shared folder or network drive and place `manifest.xml` (or `manifest.json` if using M365 previews) inside.
  2. In Excel, go to **Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs**.
  3. Add the folder path and check "Show in Menu".
  4. Go to **Insert > Get Add-ins > Shared Folder** and click on **FormulaBridge**.
- **Excel Online**:
  1. Open a blank workbook in your browser.
  2. Go to **Insert > Add-ins > Manage My Add-ins > Upload My Add-in**.
  3. Select the `manifest.json` or `manifest.xml` file.

---

## 🧠 Architecture Overview

The core of FormulaBridge is its completely independent **Translation Engine** located in `src/engine/`. It does not rely on Excel's calculation engine to parse text.

1. **Tokenizer (`parser.ts`)**: Breaks a raw formula string into semantic tokens (Functions, Strings, Booleans, References).
2. **Auto-Detect (`autoDetect.ts`)**: Analyzes function tokens against the dictionary to guess the source language with a confidence score.
3. **Separator Converter (`separatorConverter.ts`)**: Safely swaps `,` for `;` and vice versa without destroying decimal numbers.
4. **Dictionary (`dictionary.json`)**: A massive, consolidated map of 494 Excel functions across 18 languages. Generated automatically from Microsoft's localization tables via a custom build script (`scripts/build-dictionary.js`).

### The Event Loop
When a user types `=SUM(A1:A10)` in a German Excel:
1. `formulaWatcher.ts` catches the `onChanged` event.
2. It detects the language is English and the target is German.
3. The engine parses the formula and maps `SUM` -> `SUMME`.
4. The watcher disables Excel events (`context.runtime.enableEvents = false`) to prevent infinite loops.
5. It writes `=SUMME(A1:A10)` back to `range.formulasLocal`.
6. Events are re-enabled.

---

## 🤝 Contributing

Since this is a Vibe Coding learning project, feel free to fork, experiment, and break things! If you want to add a new language, simply update the `lcidMap` in `scripts/build-dictionary.js` and rebuild the dictionary.

**Happy Translating!** 🌍
