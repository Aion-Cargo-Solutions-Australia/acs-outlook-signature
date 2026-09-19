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
    '<td style="padding:3px 18px 3px 0;font-family:' + cfg.fontFamily + ";font-size:13px;font-weight:bold;color:" + c.label +
    ';white-space:nowrap;vertical-align:top;">' + label + "</td>" +
    '<td style="padding:3px 0;font-family:' + cfg.fontFamily + ";font-size:13px;color:" + c.muted + ';vertical-align:top;">' +
    valueHtml + "</td>" +
    "</tr>"
  );
}

function link(cfg, href, text) {
  return '<a href="' + escapeHtml(href) + '" style="color:' + cfg.colors.link + ';text-decoration:underline;">' + escapeHtml(text) + "</a>";
}

/**
 * @param {object} p   标准化后的用户资料（见 profile.js normalizeProfile）
 * @param {object} cfg SIGNATURE_CONFIG
 * @param {(key:string)=>string} imageSrc 返回图片 src（cid:xxx 或 https 链接）
 */
export function buildFullSignature(p, cfg, imageSrc) {
  const c = cfg.colors;
  const f = cfg.fontFamily;
  const co = cfg.company;
  const rows = [];

  if (p.mobile) rows.push(row(cfg, "Mob", '<a href="' + telHref(p.mobile) + '" style="color:' + c.muted + ';text-decoration:none;">' + escapeHtml(p.mobile) + "</a>"));
  if (p.phone) {
    let tel = '<a href="' + telHref(p.phone) + '" style="color:' + c.muted + ';text-decoration:none;">' + escapeHtml(p.phone) + "</a>";
    if (p.ext) tel += ' <span style="color:' + c.label + ';font-weight:bold;">EXT</span> ' + escapeHtml(p.ext);
    rows.push(row(cfg, "Tel", tel));
  }
  if (p.email) rows.push(row(cfg, "Email", link(cfg, "mailto:" + p.email, p.email)));
  if (co.website) rows.push(row(cfg, "Web", link(cfg, co.websiteUrl || "https://" + co.website, co.website)));
  const address = p.address || co.address;
  if (address) rows.push(row(cfg, "Address", co.addressUrl && !p.address ? link(cfg, co.addressUrl, address) : escapeHtml(address)));

  const title = cfg.showTitleInFull && p.jobTitle
    ? '<div style="font-family:' + f + ";font-size:13px;color:" + c.muted + ';padding:0 0 6px 0;">' + escapeHtml(p.jobTitle) + "</div>"
    : "";

  const logo = cfg.images.logo;
  const badge = cfg.images.badge;

  let html = "";
  html += '<div id="acs-signature" data-acs-sig="full">';
  html += signOffHtml(cfg);
  html += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;">';
  html += "<tr>";
  if (logo) {
    html += '<td width="' + (logo.width + 24) + '" style="padding:0 24px 0 0;vertical-align:middle;">' + imgTag(logo, imageSrc("logo")) + "</td>";
  }
  html += '<td style="border-left:4px solid ' + c.divider + ';padding:2px 0 2px 18px;vertical-align:top;">';
  html += '<div style="font-family:' + f + ";font-size:22px;line-height:28px;color:" + c.text + ';padding:0 0 6px 0;">' + escapeHtml(p.displayName) + "</div>";
  html += title;
  html += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;">' + rows.join("") + "</table>";
  html += "</td></tr></table>";

  if (co.disclaimer) {
    html += '<p style="margin:12px 0 0 0;font-family:' + f + ";font-size:12px;line-height:18px;color:" + c.muted + ';">' + escapeHtml(co.disclaimer);
    if (co.termsUrl) html += "<br/>" + link(cfg, co.termsUrl, co.termsText || co.termsUrl);
    html += "</p>";
  }
  if (badge) {
    html += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="margin-top:12px;"><tr><td>' + imgTag(badge, imageSrc("badge")) + "</td></tr></table>";
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
  if (p.phone) parts.push('<b style="color:' + c.label + ';">T</b>&nbsp;' + escapeHtml(p.phone) + (p.ext ? "&nbsp;ext&nbsp;" + escapeHtml(p.ext) : ""));
  if (co.website) parts.push(link(cfg, co.websiteUrl || "https://" + co.website, co.website));

  let html = '<div id="acs-signature" data-acs-sig="short">';
  if (cfg.signOffInShort) html += signOffHtml(cfg);
  html += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;"><tr>';
  html += '<td style="border-left:3px solid ' + c.accent + ';padding:1px 0 1px 10px;font-family:' + f + ";font-size:12px;line-height:19px;color:" + c.muted + ';">';
  html += line1 + "<br/>" + line2 + "<br/>" + parts.join(sep);
  html += "</td></tr></table></div>";
  return html;
}
