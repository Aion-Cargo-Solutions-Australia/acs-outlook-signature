/**
 * 纯函数：根据用户资料 + 配置生成签名 HTML。不依赖 Office.js，可在 Node 中测试。
 * Outlook 兼容要点：table 布局、内联样式、img 显式 width/height、不用 SVG / CSS class。
 */

export function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function telHref(phone) {
  return "tel:" + String(phone).replace(/[^\d+]/g, "");
}

function signOffHtml(cfg) {
  if (!cfg.signOff) return "";
  return (
    '<p style="margin:0 0 10px 0;font-family:' + cfg.fontFamily + ";font-size:14px;color:" + cfg.colors.text + ';">' +
    escapeHtml(cfg.signOff) +
    "</p>"
  );
}

function imgTag(img, src) {
  return (
    '<img src="' + escapeHtml(src) + '" width="' + img.width + '" height="' + img.height + '" alt="' + escapeHtml(img.alt) +
    '" style="display:block;width:' + img.width + "px;height:" + img.height + 'px;border:0;outline:none;" />'
  );
}

function row(cfg, label, valueHtml) {
  const c = cfg.colors;
  return (
    "<tr>" +
    '<td width="16" style="padding:2px 8px 2px 0;font-family:' + cfg.fontFamily + ";font-size:11px;line-height:18px;font-weight:bold;color:" + c.label +
    ';white-space:nowrap;vertical-align:top;">' + label + "</td>" +
    '<td style="padding:2px 0;font-family:' + cfg.fontFamily + ";font-size:12px;line-height:18px;color:" + c.muted + ';vertical-align:top;">' +
    valueHtml + "</td>" +
    "</tr>"
  );
}

function link(cfg, href, text, color, underline) {
  return '<a href="' + escapeHtml(href) + '" style="color:' + (color || cfg.colors.link) + ";text-decoration:" + (underline ? "underline" : "none") + ';">' + escapeHtml(text) + "</a>";
}

/**
 * 完整签名
 *  ┌──────────┬─────────────────────────────┐
 *  │ ACS Logo │ Name                        │
 *  │          │ Title                       │
 *  │          │ M / T / E / W / A           │
 *  └──────────┴─────────────────────────────┘
 *  ───────────── 细分隔线 ─────────────
 *  This email is subject to our Confidentiality Statement & Terms and Conditions.
 *  [WCA 徽章]
 *
 * @param {object} p   标准化后的用户资料（见 profile.js normalizeProfile）
 * @param {object} cfg SIGNATURE_CONFIG
 * @param {(key:string)=>string} imageSrc 返回图片 src（cid:xxx 或 https 链接）
 */
export function buildFullSignature(p, cfg, imageSrc) {
  const c = cfg.colors;
  const f = cfg.fontFamily;
  const co = cfg.company;
  const plain = (t) => escapeHtml(t);
  const rows = [];

  if (p.mobile) rows.push(row(cfg, "M", link(cfg, telHref(p.mobile), p.mobile, c.muted)));
  if (p.phone || p.ext) {
    let tel = p.phone ? link(cfg, telHref(p.phone), p.phone, c.muted) : "";
    if (p.ext) tel += (tel ? "&nbsp;&nbsp;" : "") + '<span style="color:' + c.text + ';">EXT ' + plain(p.ext) + "</span>";
    rows.push(row(cfg, "T", tel));
  }
  if (p.email) rows.push(row(cfg, "E", link(cfg, "mailto:" + p.email, p.email, c.muted)));
  if (co.website) rows.push(row(cfg, "W", link(cfg, co.websiteUrl || "https://" + co.website, co.website, c.muted)));
  const address = p.address || co.address;
  if (address) rows.push(row(cfg, "A", co.addressUrl && !p.address ? link(cfg, co.addressUrl, address, c.muted) : plain(address)));

  const subtitle = [cfg.showTitleInFull && p.jobTitle ? p.jobTitle : ""].filter(Boolean);
  const logo = cfg.images.logo;
  const badge = cfg.images.badge;
  const leftW = logo ? logo.width : 0;

  let html = '<div id="acs-signature" data-acs-sig="full">';
  html += signOffHtml(cfg);
  html += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;">';
  html += "<tr>";

  // 左列：Logo + 徽章
  if (leftW) {
    html += '<td width="' + leftW + '" valign="top" style="padding:2px 18px 0 0;vertical-align:top;">';
    html += imgTag(logo, imageSrc("logo"));
    html += "</td>";
  }

  // 右列：姓名 / 职位 / 联系方式
  html += '<td valign="top" style="border-left:2px solid ' + c.divider + ';padding:0 0 0 18px;vertical-align:top;">';
  html += '<div style="font-family:' + f + ";font-size:20px;line-height:24px;font-weight:bold;color:" + c.text + ';letter-spacing:0.2px;">' + plain(p.displayName) + "</div>";
  if (subtitle.length) {
    html += '<div style="font-family:' + f + ";font-size:12px;line-height:18px;color:" + c.muted + ';padding:2px 0 10px 0;">' +
      subtitle.map(plain).join('<span style="color:' + c.accent + ';">&nbsp;&nbsp;|&nbsp;&nbsp;</span>') + "</div>";
  } else {
    html += '<div style="height:10px;line-height:10px;font-size:1px;">&nbsp;</div>';
  }
  html += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;">' + rows.join("") + "</table>";
  html += "</td></tr></table>";

  // 底部：细线 + 一行免责声明
  if (co.termsUrl || co.disclaimerPrefix) {
    html += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;margin-top:14px;"><tr>';
    html += '<td style="border-top:1px solid ' + c.hairline + ';padding:8px 0 0 0;font-family:' + f + ";font-size:10.5px;line-height:15px;color:#9a9a9a;\">";
    html += plain(co.disclaimerPrefix || "") + (co.disclaimerPrefix ? " " : "");
    html += co.termsUrl ? link(cfg, co.termsUrl, co.termsText || co.termsUrl, "#9a9a9a", true) : "";
    html += ".</td></tr></table>";
  }
  // 最下方：WCA 徽章（比 ACS Logo 小）
  if (badge) {
    html += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;margin-top:10px;"><tr><td>' +
      imgTag(badge, imageSrc("badge")) + "</td></tr></table>";
  }
  html += "</div>";
  return html;
}

/**
 * 精简签名：无图片，三行。用于同一邮件会话中本人第二次及以后的回复。
 */
export function buildShortSignature(p, cfg) {
  const c = cfg.colors;
  const f = cfg.fontFamily;
  const co = cfg.company;
  const sep = '<span style="color:#c8c8c8;">&nbsp;|&nbsp;</span>';

  const line1 = '<b style="font-size:14px;color:' + c.text + ';">' + escapeHtml(p.displayName) + "</b>" +
    (cfg.showTitleInShort && p.jobTitle ? sep + '<span style="color:' + c.muted + ';">' + escapeHtml(p.jobTitle) + "</span>" : "");

  const addr = p.address || co.address;
  const line2 = '<b style="color:' + c.accent + ';">' + escapeHtml(co.shortName) + "</b>&nbsp;" +
    '<span style="color:' + c.text + ';">' + escapeHtml(co.name) + "</span>" +
    (cfg.showAddressInShort && addr ? sep + '<span style="color:' + c.muted + ';">' + escapeHtml(addr) + "</span>" : "");

  const parts = [];
  if (p.mobile) parts.push('<b style="color:' + c.label + ';">M</b>&nbsp;' + escapeHtml(p.mobile));
  if (p.phone) parts.push('<b style="color:' + c.label + ';">T</b>&nbsp;' + escapeHtml(p.phone) + (p.ext ? "&nbsp;EXT&nbsp;" + escapeHtml(p.ext) : ""));
  if (co.website) parts.push(link(cfg, co.websiteUrl || "https://" + co.website, co.website));

  let html = '<div id="acs-signature" data-acs-sig="short">';
  if (cfg.signOffInShort) html += signOffHtml(cfg);
  html += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;"><tr>';
  html += '<td style="border-left:3px solid ' + c.accent + ';padding:1px 0 1px 10px;font-family:' + f + ";font-size:12px;line-height:19px;color:" + c.muted + ';">';
  html += line1 + "<br/>" + line2 + "<br/>" + parts.join(sep);
  html += "</td></tr></table></div>";
  return html;
}
