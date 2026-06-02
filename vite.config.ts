import { defineConfig } from "vite";
import { readFileSync } from "fs";
import { resolve } from "path";

// Office Add-ins require HTTPS even in dev mode
let httpsConfig: { key?: Buffer; cert?: Buffer } | false = false;

try {
  // Use office-addin-dev-certs generated certificates
  const certDir = resolve(
    process.env.HOME || process.env.USERPROFILE || "",
    ".office-addin-dev-certs"
  );
  httpsConfig = {
    key: readFileSync(resolve(certDir, "localhost.key")),
    cert: readFileSync(resolve(certDir, "localhost.crt")),
  };
} catch {
  console.warn(
    "⚠️  No HTTPS certs found. Run: npx office-addin-dev-certs install"
  );
  console.warn("   Falling back to HTTP (sideloading may not work).");
}

export default defineConfig({
  root: ".",
  publicDir: "assets",
  build: {
    outDir: "dist",
    sourcemap: true,
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "index.html"),
      },
    },
  },
  server: {
    https: httpsConfig || undefined,
    port: 3000,
    headers: {
      // Required for Office Add-in iframe embedding
      "Access-Control-Allow-Origin": "*",
    },
  },
  resolve: {
    alias: {
      "@": resolve(__dirname, "src"),
    },
  },
});
