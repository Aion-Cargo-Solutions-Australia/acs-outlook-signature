// 构建：生成配置模块 → esbuild 打包 → 拷贝静态文件 → 生成 manifest.xml
import { build } from "esbuild";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const dist = path.join(root, "dist");
const gen = path.join(root, "src", "generated");
const cfg = JSON.parse(fs.readFileSync(path.join(root, "addin.config.json"), "utf8"));
const { SIGNATURE_CONFIG } = await import(pathToFileURL(path.join(root, "src", "config.js")).href);

// ---- 校验 ----
const baseUrl = String(process.env.ADDIN_BASE_URL || cfg.baseUrl).replace(/\/+$/, "");
const clientId = process.env.ADDIN_CLIENT_ID || cfg.clientId;
const tenantId = process.env.ADDIN_TENANT_ID || cfg.tenantId || "organizations";
if (!/^https:\/\//.test(baseUrl)) throw new Error("baseUrl 必须是 https:// 开头");
const u = new URL(baseUrl);
if (u.pathname && u.pathname !== "/") {
  console.warn(
    "\n⚠️  baseUrl 带有子路径 (" + u.pathname + ")。经典 Outlook(Windows) 需要 " + u.origin +
      "/.well-known/microsoft-officeaddins-allowed.json 位于域名根目录，建议使用自定义子域名（如 https://sig.breezefreight.com）。\n"
  );
}
if (/^0{8}-/.test(clientId)) console.warn("⚠️  clientId 还是占位符，请在 addin.config.json 中填入 Entra 应用的 Application (client) ID");

// ---- 生成模块 ----
fs.mkdirSync(gen, { recursive: true });
fs.writeFileSync(
  path.join(gen, "addin.js"),
  "// 自动生成，勿改\nexport const ADDIN = " + JSON.stringify({ baseUrl, clientId, tenantId, version: cfg.version }, null, 2) + ";\n"
);
const images = {};
for (const [key, img] of Object.entries(SIGNATURE_CONFIG.images)) {
  if (!img) continue;
  const file = path.join(root, "src", "assets", img.file);
  const b64 = fs.readFileSync(file).toString("base64");
  images[key] = { file: img.file, cid: img.file, base64: b64 };
  if (b64.length > 200000) console.warn("⚠️  " + img.file + " 较大（" + Math.round(b64.length / 1024) + "KB base64），建议压缩");
}
fs.writeFileSync(path.join(gen, "images.js"), "// 自动生成，勿改\nexport const IMAGES = " + JSON.stringify(images) + ";\n");

// ---- 打包 ----
fs.rmSync(dist, { recursive: true, force: true });
fs.mkdirSync(path.join(dist, "assets"), { recursive: true });
const common = { bundle: true, format: "iife", target: ["es2016"], minify: true, sourcemap: false, legalComments: "none", logLevel: "warning" };
await build({ ...common, entryPoints: [path.join(root, "src/launchevent/launchevent.js")], outfile: path.join(dist, "launchevent.js") });
await build({ ...common, entryPoints: [path.join(root, "src/taskpane/taskpane.js")], outfile: path.join(dist, "taskpane.js") });

// ---- 静态文件 ----
const sub = (s) =>
  s.replaceAll("{{BASE_URL}}", baseUrl)
    .replaceAll("{{ADDIN_ID}}", cfg.addinId)
    .replaceAll("{{VERSION}}", cfg.version)
    .replaceAll("{{PROVIDER}}", cfg.providerName)
    .replaceAll("{{DISPLAY_NAME}}", cfg.displayName)
    .replaceAll("{{SUPPORT_URL}}", SIGNATURE_CONFIG.company.websiteUrl || baseUrl);
fs.copyFileSync(path.join(root, "src/launchevent/commands.html"), path.join(dist, "commands.html"));
fs.copyFileSync(path.join(root, "src/taskpane/taskpane.html"), path.join(dist, "taskpane.html"));
fs.copyFileSync(path.join(root, "src/taskpane/taskpane.css"), path.join(dist, "taskpane.css"));
for (const f of fs.readdirSync(path.join(root, "src/assets"))) fs.copyFileSync(path.join(root, "src/assets", f), path.join(dist, "assets", f));
fs.mkdirSync(path.join(dist, ".well-known"), { recursive: true });
fs.writeFileSync(
  path.join(dist, ".well-known/microsoft-officeaddins-allowed.json"),
  sub(fs.readFileSync(path.join(root, "src/well-known/microsoft-officeaddins-allowed.json"), "utf8"))
);
fs.writeFileSync(path.join(dist, "manifest.xml"), sub(fs.readFileSync(path.join(root, "manifest.template.xml"), "utf8")));
fs.writeFileSync(path.join(dist, ".nojekyll"), "");
if (u.hostname && !u.hostname.endsWith("github.io")) fs.writeFileSync(path.join(dist, "CNAME"), u.hostname + "\n");
fs.writeFileSync(path.join(dist, "index.html"), '<!doctype html><meta charset="utf-8"><title>ACS Signature</title><p>ACS Outlook signature add-in. <a href="manifest.xml">manifest.xml</a></p>');

console.log("✅ build 完成 → dist/  (baseUrl=" + baseUrl + ")");
