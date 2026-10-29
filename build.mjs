// Arma la carpeta www/ (lo que va dentro del APK) a partir de web/.
// Empaqueta Firebase, jsPDF y la tipografía para que la app funcione sin señal desde el primer uso.
import { build } from "esbuild";
import { cpSync, rmSync, mkdirSync } from "node:fs";

rmSync("www", { recursive: true, force: true });
mkdirSync("www");
for (const f of ["index.html", "styles.css", "manifest.webmanifest", "sw.js"]) cpSync(`web/${f}`, `www/${f}`);
cpSync("web/icons", "www/icons", { recursive: true });

await build({
  entryPoints: { app: "web/app.js" },
  bundle: true,
  format: "esm",
  target: ["chrome80", "safari15"],
  minify: true,
  outdir: "www",
  loader: { ".woff": "file", ".woff2": "file" },
  assetNames: "fonts/[name]-[hash]",
  logLevel: "info"
});
console.log("www/ listo");
