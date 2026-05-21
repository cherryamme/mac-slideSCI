import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const projectRoot = resolve(__dirname, "..");
const manifestPath = resolve(projectRoot, "manifest.xml");
const distDir = resolve(projectRoot, "dist");
const outputPath = resolve(distDir, "manifest.xml");
const baseUrl = (process.env.PUBLIC_BASE_URL ?? "").replace(/\/+$/u, "");

if (!baseUrl) {
  console.error("[build-manifest] 缺少 PUBLIC_BASE_URL 环境变量，例如 https://cherryamme.github.io/mac-slideSCI");
  process.exit(1);
}

const source = readFileSync(manifestPath, "utf8");
const rewritten = source.replaceAll("https://localhost:3000", baseUrl);

mkdirSync(distDir, { recursive: true });
writeFileSync(outputPath, rewritten);
writeFileSync(resolve(distDir, ".nojekyll"), "");
console.log(`[build-manifest] Wrote ${outputPath} (base=${baseUrl})`);
