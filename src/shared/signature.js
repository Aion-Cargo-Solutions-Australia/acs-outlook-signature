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
    '<td width="54" style="width:54px;padding:1px 10px 1px 0;font-family:' + cfg.fontFamily + ";font-size:11px;line-height:18px;font-weight:bold;color:" + c.label +
    ';white-space:nowrap;vertical-align:top;">' + escapeHtml(label) + "</td>" +
    '<td style="padding:1px 0;font-family:' + cfg.fontFamily + ";font-size:12px;line-height:18px;color:" + c.value + ';white-space:nowrap;vertical-align:top;">' +
    valueHtml + "</td>" +
    "</tr>"
  );
}

function extHtml(cfg, ext) {
  return '<span style="color:' + cfg.colors.muted + ';">Ext.</span>&nbsp;' + escapeHtml(ext);
}

function link(cfg, href, text, color, underline) {
  return '<a href="' + escapeHtml(href) + '" style="color:' + (color || cfg.colors.link) + ";text-decoration:" + (underline ? "underline" : "none") + ';">' + escapeHtml(text) + "</a>";
}

/**
 * 完整签名
 *  ┌────────────┬──────────────────────────────┐
 *  │            │ Name                         │
 *  │  ACS Logo  │ Title                        │
 *  │ (与右侧等高)│ Mobile / Tel / Email / Web / │
 *  │            │ Address                      │
 *  ├────────────┴──────────────────────────────┤  ← 细分隔线（与上方同宽）
 *  │ [WCA]  This email is subject to our Confidentiality Statement & T&C. │
 *  └───────────────────────────────────────────┘
 *
 * @param {object} p   标准化后的用户资料（见 profile.js normalizeProfile）
 * @param {object} cfg SIGNATURE_CONFIG
 * @param {(key:string)=>string} imageSrc 返回图片 src（cid:xxx 或 https 链接）
 */
export function buildFullSignature(p, cfg, imageSrc) {
  const c = cfg.colors;
  const f = cfg.fontFamily;
  const co = cfg.company;
  const L = cfg.labels;
  const plain = (t) => escapeHtml(t);
  const rows = [];

  if (p.mobile) rows.push(row(cfg, L.mobile, link(cfg, telHref(p.mobile), p.mobile, c.value)));
  if (p.phone || p.ext) {
    let tel = p.phone ? link(cfg, telHref(p.phone), p.phone, c.value) : "";
    if (p.ext) tel += (tel ? "&nbsp;&nbsp;" : "") + extHtml(cfg, p.ext);
    rows.push(row(cfg, L.phone, tel));
  }
  if (p.email) rows.push(row(cfg, L.email, link(cfg, "mailto:" + p.email, p.email, c.value)));
  if (co.website) rows.push(row(cfg, L.web, link(cfg, co.websiteUrl || "https://" + co.website, co.website, c.value)));
  // 地址统一使用公司地址（不读个人资料里的地址，旧版本缓存里可能存有员工住址）
  if (co.address) rows.push(row(cfg, L.address, co.addressUrl ? link(cfg, co.addressUrl, co.address, c.value) : plain(co.address)));

  const logo = cfg.images.logo;
  const badge = cfg.images.badge;
  const hasDisclaimer = !!(co.termsUrl || co.disclaimerPrefix);

  let html = '<div id="acs-signature" data-acs-sig="full">';
  html += signOffHtml(cfg);
  html += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;">';
  html += "<tr>";

  // 左列：Logo，高度与右侧信息块一致，垂直居中
  if (logo) {
    html += '<td width="' + logo.width + '" valign="middle" style="padding:0 20px 0 0;vertical-align:middle;">' + imgTag(logo, imageSrc("logo")) + "</td>";
  }

  // 右列：姓名 / 职位 / 联系方式
  html += '<td valign="middle" style="border-left:2px solid ' + c.divider + ';padding:2px 0 2px 20px;vertical-align:middle;">';
  html += '<div style="font-family:' + f + ";font-size:20px;line-height:24px;font-weight:bold;color:" + c.text + ';">' + plain(p.displayName) + "</div>";
  if (cfg.showTitleInFull && p.jobTitle) {
    html += '<div style="font-family:' + f + ";font-size:12px;line-height:18px;color:" + c.muted + ';padding:1px 0 0 0;">' + plain(p.jobTitle) + "</div>";
  }
  html += '<div style="height:10px;line-height:10px;font-size:1px;">&nbsp;</div>';
  html += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;">' + rows.join("") + "</table>";
  html += "</td></tr>";

  // 底部（跨两列，与上方同宽）：细分隔线 + WCA 徽章 + 一行免责声明
  if (hasDisclaimer || badge) {
    html += '<tr><td colspan="' + (logo ? 2 : 1) + '" style="padding:14px 0 0 0;">';
    html += '<table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;"><tr>';
    html += '<td style="border-top:1px solid ' + c.hairline + ';padding:10px 0 0 0;">';
    html += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;"><tr>';
    if (badge) {
      html += '<td valign="middle" style="padding:0 14px 0 0;vertical-align:middle;">' + imgTag(badge, imageSrc("badge")) + "</td>";
    }
    if (hasDisclaimer) {
      html += '<td valign="middle" style="font-family:' + f + ";font-size:10.5px;line-height:15px;color:" + c.footer + ';vertical-align:middle;">';
      html += plain(co.disclaimerPrefix || "") + (co.disclaimerPrefix ? " " : "");
      html += co.termsUrl ? link(cfg, co.termsUrl, co.termsText || co.termsUrl, c.footer, true) : "";
      html += ".</td>";
    }
    html += "</tr></table></td></tr></table></td></tr>";
  }
  html += "</table></div>";
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

  const addr = co.address;
  const line2 = '<b style="color:' + c.accent + ';">' + escapeHtml(co.shortName) + "</b>&nbsp;" +
    '<span style="color:' + c.text + ';">' + escapeHtml(co.name) + "</span>" +
    (cfg.showAddressInShort && addr ? sep + '<span style="color:' + c.muted + ';">' + escapeHtml(addr) + "</span>" : "");

  const L = cfg.labels;
  const parts = [];
  if (p.mobile) parts.push('<b style="color:' + c.label + ';">' + escapeHtml(L.mobile) + "</b>&nbsp;" + escapeHtml(p.mobile));
  if (p.phone) parts.push('<b style="color:' + c.label + ';">' + escapeHtml(L.phone) + "</b>&nbsp;" + escapeHtml(p.phone) + (p.ext ? "&nbsp;&nbsp;" + extHtml(cfg, p.ext) : ""));
  if (co.website) parts.push(link(cfg, co.websiteUrl || "https://" + co.website, co.website));

  let html = '<div id="acs-signature" data-acs-sig="short">';
  if (cfg.signOffInShort) html += signOffHtml(cfg);
  html += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;"><tr>';
  html += '<td style="border-left:3px solid ' + c.accent + ';padding:1px 0 1px 10px;font-family:' + f + ";font-size:12px;line-height:19px;color:" + c.muted + ';">';
  html += line1 + "<br/>" + line2 + "<br/>" + parts.join(sep);
  html += "</td></tr></table></div>";
  return html;
}
