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

## 📦 Deployment & Sideloading (No Store Registration Needed!)

FormulaBridge is hosted live at **https://formula-bridge.vercel.app**. Since it is an open-source project, you can use it without downloading any development tools or paying any Microsoft developer fees.

### For Average Users (Store-Free Installation)
1. Open [formula-bridge.vercel.app](https://formula-bridge.vercel.app) in your web browser.
2. Click the **Download manifest.xml** button to save the manifest to your device.
3. Open Excel (Desktop or Web), go to the **Insert** tab, click **Add-ins** > **Manage My Add-ins** > **Upload My Add-in**, and select the downloaded `manifest.xml` file.
4. The FormulaBridge 🌉 button will appear in your Home tab ribbon!

---

## 🛠️ Local Development Setup

If you want to run the project locally and modify the code:

### 1. Install Dependencies
```bash
npm install
```

### 2. Generate Development Certificates
Office Add-ins require HTTPS for security. Generate local certificates:
```bash
npx office-addin-dev-certs install
```

### 3. Manifest Files
- **`manifest.local.xml`**: Used for local development and testing. Points to `https://localhost:3000`.
- **`manifest.xml`**: Used for production deployment. Points to the live `https://formula-bridge.vercel.app`.

### 4. Start the Dev Server & Sideload
Start the Vite server:
```bash
npm run dev
```
Then, sideload the local version into Excel automatically:
```bash
npm run sideload
```
*(This command automatically runs `office-addin-debugging` using `manifest.local.xml`)*

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
