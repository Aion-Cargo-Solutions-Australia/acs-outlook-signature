// 生成 preview.html：用示例资料渲染完整签名 + 精简签名（图片内嵌为 data URI），用于设计确认
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SIGNATURE_CONFIG as CFG } from "../src/config.js";
import { buildFullSignature, buildShortSignature } from "../src/shared/signature.js";
import { normalizeProfile } from "../src/shared/profile.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const me = {
  displayName: "Ivy Hu", jobTitle: "Managing Director – Australia", mail: "ivy.hu@aioncargo.com",
  mobilePhone: "+61 425 666 802", businessPhones: ["+61 2 9160 2300 EXT 601"],
};
const p = normalizeProfile(me, CFG, {});
const dataUri = (k) => "data:image/png;base64," + fs.readFileSync(path.join(root, "src/assets", CFG.images[k].file)).toString("base64");
const full = buildFullSignature(p, CFG, dataUri);
const short = buildShortSignature(p, CFG, dataUri);

const quoted = (inner) =>
  '<div style="border-top:1px solid #e1e1e1;margin-top:18px;padding-top:10px;font:12px Arial;color:#555;"><b>From:</b> John Smith &lt;john@client.com&gt;<br/><b>Sent:</b> Friday, 18 September 2026 3:12 PM<br/><b>Subject:</b> RE: Shipment quote SYD → SIN</div>' + (inner || "");

const html = `<!doctype html><html><head><meta charset="utf-8"><title>ACS 签名预览</title>
<style>body{margin:0;background:#eef0f3;font:14px "Segoe UI",Arial;color:#222}.wrap{max-width:900px;margin:0 auto;padding:24px}
h1{font-size:20px;margin:0 0 4px}.sub{color:#666;margin:0 0 20px}.mail{background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);padding:22px 26px;margin-bottom:22px}
.tag{display:inline-block;font-size:12px;font-weight:600;padding:2px 8px;border-radius:10px;margin-bottom:12px}.full{background:#fde7ed;color:#b00c38}.short{background:#e8f1fd;color:#0b5cad}
.body{font:14px Arial;margin:0 0 16px}
.dark{background:#262626}.dark .body,.dark #acs-signature [style*="#1a1a1a"],.dark #acs-signature [style*="#333333"]{color:#e3e3e3!important}
.dark #acs-signature [style*="#6b6b6b"],.dark #acs-signature [style*="#8c8c8c"]{color:#a8a8a8!important}.dark .tag{background:#3a3a3a;color:#ddd}</style></head><body><div class="wrap">
<h1>ACS 邮件签名 — 设计预览</h1><p class="sub">示例数据：Ivy Hu。实际签名中的个人信息从 Entra ID 自动读取。</p>
<div class="mail"><span class="tag full">① 新邮件 / 会话中第一次回复 / 转发 → 完整签名</span><p class="body">Hi John,<br/><br/>Please find the updated quote attached.</p>${full}</div>
<div class="mail dark"><span class="tag">③ 暗色模式（模拟 Outlook 深色主题：文字颜色被反转，图片不变）</span><p class="body">Hi John,<br/><br/>Please find the updated quote attached.</p>${full}</div>
<div class="mail"><span class="tag short">② 同一会话中后续回复 → 精简签名（无图片）</span><p class="body">Hi John,<br/><br/>Confirmed — booking for Monday's vessel.</p>${short}${quoted()}</div>
</div></body></html>`;
fs.writeFileSync(path.join(root, "preview.html"), html);
console.log("✅ preview.html 已生成");
