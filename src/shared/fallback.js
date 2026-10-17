/**
 * 兜底签名：加载项没运行的场景（Outlook 手机版、单独打开的共享邮箱等）。纯函数，不依赖 Office.js，可在 Node 中测试。
 *  - 服务器端：Exchange 邮件流规则在共享邮箱发出的邮件末尾追加团队信息（renderRulesScript 生成规则脚本）
 *  - 手机端：员工在 Outlook 手机版里给共享邮箱账户设置的文字签名（renderMobilePage 生成可复制的页面）
 * 两者都含同一行“团队名 | 公司名”（markerLine）。规则看到这一行，或看到加载项打的邮件头，就不再追加，避免重复和叠加。
 */
import { activeBanner, escapeHtml } from "./signature.js";
import { SIGNATURE_HEADER } from "./header.js";

export function markerLine(box, cfg) {
  return box.team + " | " + cfg.company.name;
}

/** 把一段文字变成规则用的正则（空白宽松匹配，Exchange 的规则匹配本身不区分大小写） */
function phrasePattern(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/ +/g, "\\s+");
}

export function markerPattern(box, cfg) {
  return phrasePattern(markerLine(box, cfg));
}

/**
 * 节日横幅期间，加载项没运行的邮件（手机、单独打开的共享邮箱）放不了横幅图片，改由规则追加文字版放假安排（横幅的 alt 文字）。
 * 横幅已过期则返回 ""，规则脚本里也就没有这条规则。
 */
export function buildHolidayHtml(cfg) {
  const b = activeBanner(cfg);
  return b ? '<div style="margin-top:10px;font-family:' + cfg.fontFamily + ";font-size:12px;line-height:18px;color:" + cfg.colors.accent +
    ';font-weight:bold;">' + escapeHtml(b.alt) + "</div>" : "";
}

// 去重只看开头一段（“Merry Christmas & Happy New Year. Office closed:”），整句正则太长
export function holidayPattern(cfg) {
  return phrasePattern(activeBanner(cfg).alt.slice(0, 48).trim());
}

function telLine(box) {
  const ext = String(box.ext || "").replace(/^\s*(ext\.?|x)\s*/i, "").trim();
  return "Tel " + box.phone + (ext ? " Ext. " + ext : "");
}

/** Exchange 追加在邮件末尾的团队信息（HTML，Exchange 限制 5000 字符；不放图片，收件人默认也不会加载外链图片） */
export function buildFallbackHtml(box, cfg) {
  const c = cfg.colors;
  const co = cfg.company;
  const email = box.addresses[0];
  const a = (href, text, color, underline) =>
    '<a href="' + escapeHtml(href) + '" style="color:' + color + ";text-decoration:" + (underline ? "underline" : "none") + ';">' + escapeHtml(text) + "</a>";
  return (
    '<div style="margin-top:16px;padding-top:8px;border-top:1px solid ' + c.hairline + ";font-family:" + cfg.fontFamily +
    ";font-size:12px;line-height:18px;color:" + c.muted + ';">' +
    '<b style="color:' + c.accent + ';">' + escapeHtml(box.team) + "</b> | " + escapeHtml(co.name) + "<br>" +
    escapeHtml(telLine(box)) + " | " + a("mailto:" + email, email, c.muted) + " | " + a(co.websiteUrl, co.website, c.muted) + "<br>" +
    escapeHtml(co.address) + "<br>" +
    '<span style="font-size:10px;color:' + c.footer + ';">' + escapeHtml(co.disclaimerPrefix) + " " +
    a(co.termsUrl, co.termsText, c.footer, true) + ".</span>" +
    "</div>"
  );
}

/** Outlook 手机版里给共享邮箱账户设置的文字签名；[Your name] 由员工换成自己的名字 */
export function mobileSignatureText(box, cfg) {
  const co = cfg.company;
  return [
    cfg.signOff,
    "[Your name]",
    markerLine(box, cfg),
    telLine(box) + " | " + box.addresses[0] + " | " + co.website,
    co.address,
  ].filter(Boolean).join("\n");
}

function ruleName(box) {
  return "ACS signature fallback - " + box.team;
}

/**
 * 生成 Exchange Online PowerShell 脚本：每个共享邮箱一条邮件流规则。可以重复运行（已存在的规则会被更新）。
 * 脚本只用 ASCII，Windows PowerShell 5.1 按本地编码读取也不会乱码。
 */
export function renderRulesScript(cfg) {
  const rule = (name, from, pattern, html, expiry) => [
    "  @{",
    '    Name    = "' + name + '"',
    "    From    = @(" + from.map((a) => '"' + a + '"').join(", ") + ")",
    "    Pattern = '" + pattern + "'",
  ].concat(expiry ? ['    Expiry  = [datetime]"' + expiry + '"'] : [], ["    Html    = @'", html, "'@", "  }"]).join("\n");
  const rules = cfg.sharedMailboxes.map((box) => rule(ruleName(box), [box.addresses[0]], markerPattern(box, cfg), buildFallbackHtml(box, cfg)));
  // 放在最后：新建的规则优先级排在后面，文字追加在团队信息下面；到横幅的 until 自动失效
  const banner = activeBanner(cfg);
  if (banner) {
    rules.push(rule("ACS signature fallback - holiday notice", cfg.sharedMailboxes.map((box) => box.addresses[0]), holidayPattern(cfg),
      buildHolidayHtml(cfg), new Date(banner.until).toISOString()));
  }

  return [
    "# Generated from src/config.js (sharedMailboxes) by the build. Do not edit by hand.",
    "#",
    "# Exchange Online fallback signature for the ACS shared mailboxes.",
    "# When a message is sent from one of these mailboxes and the ACS Outlook add-in did not insert",
    "# a signature (Outlook mobile, a shared mailbox opened on its own, ...), Exchange appends the",
    "# team contact block at the very end of the message.",
    "#",
    "# Skipped when:",
    "#   - the add-in inserted a signature (header " + SIGNATURE_HEADER + " contains \"inserted\"), or",
    "#   - the message already contains the team line \"<team> | " + cfg.company.name + "\"",
    "#     (Outlook mobile signature, or a fallback block added earlier in the same thread).",
    "# While the seasonal banner is on, a last rule appends the holiday notice as text (no image possible",
    "# here). It has an ExpiryDate and switches itself off; the Outlook mobile signature does not skip it.",
    "#",
    "# Usage (Exchange administrator):",
    "#   Install-Module ExchangeOnlineManagement -Scope CurrentUser   # first time only",
    "#   Connect-ExchangeOnline",
    "#   .\\fallback-rules.ps1",
    "# Safe to run again: existing rules are updated in place.",
    "",
    '$ErrorActionPreference = "Stop"',
    "",
    "$rules = @(",
    rules.join("\n"),
    ")",
    "",
    "foreach ($r in $rules) {",
    "  $params = @{",
    "    From                                 = $r.From",
    '    ApplyHtmlDisclaimerLocation          = "Append"',
    "    ApplyHtmlDisclaimerText              = $r.Html",
    '    ApplyHtmlDisclaimerFallbackAction    = "Ignore"',
    '    ExceptIfHeaderContainsMessageHeader  = "' + SIGNATURE_HEADER + '"',
    '    ExceptIfHeaderContainsWords          = "inserted"',
    "    ExceptIfSubjectOrBodyMatchesPatterns = $r.Pattern",
    '    Mode                                 = "Enforce"',
    '    Comments                             = "Generated by the ACS Outlook signature project (fallback-rules.ps1)."',
    "  }",
    "  if ($r.Expiry) { $params.ExpiryDate = $r.Expiry }",
    "  if (Get-TransportRule -Identity $r.Name -ErrorAction SilentlyContinue) {",
    "    Set-TransportRule -Identity $r.Name @params",
    '    Write-Host "Updated: $($r.Name)"',
    "  } else {",
    "    New-TransportRule -Name $r.Name @params | Out-Null",
    '    Write-Host "Created: $($r.Name)"',
    "  }",
    "}",
    "",
  ].join("\n");
}

/** 手机签名页面：员工在手机上打开，长按复制自己团队的那一段，粘贴到 Outlook 手机版的签名设置里 */
export function renderMobilePage(cfg) {
  const blocks = cfg.sharedMailboxes
    .map((box) =>
      '<section><h2>' + escapeHtml(box.team) + '</h2><p class="addr">' + escapeHtml(box.addresses[0]) + "</p><pre>" +
      escapeHtml(mobileSignatureText(box, cfg)) + "</pre></section>"
    )
    .join("\n");
  return (
    '<!doctype html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    "<title>ACS 共享邮箱手机签名</title><style>" +
    "body{margin:0;padding:16px;font:15px/1.5 -apple-system,'Segoe UI',Arial,sans-serif;color:#1a1a1a;background:#f4f5f7}" +
    "h1{font-size:19px;margin:0 0 8px}ol{padding-left:20px;margin:0 0 16px}section{background:#fff;border-radius:10px;padding:12px 14px;margin:0 0 12px}" +
    "h2{font-size:16px;margin:0;color:#D70F44}.addr{margin:0 0 8px;color:#6b6b6b;font-size:13px}" +
    "pre{white-space:pre-wrap;word-break:break-word;margin:0;font:14px/1.5 Menlo,Consolas,monospace;background:#f7f7f7;padding:10px;border-radius:6px;-webkit-user-select:all;user-select:all}" +
    "</style></head><body><h1>ACS 共享邮箱手机签名</h1><ol>" +
    "<li>长按下面自己团队的那一段，全选复制，把 <b>[Your name]</b> 换成你的名字。</li>" +
    "<li>Outlook 手机版 → 左上角头像 → 设置（齿轮）→ 签名 → 打开“按账户设置签名” → 选这个共享邮箱 → 粘贴保存。</li>" +
    "<li>不要改“团队名 | " + escapeHtml(cfg.company.name) + "”这一行：服务器靠它判断邮件已经带了签名，改了会在邮件末尾再追加一份。</li>" +
    "</ol>\n" + blocks + "\n</body></html>\n"
  );
}
