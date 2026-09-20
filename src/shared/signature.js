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

function row(cfg, label, valueHtml, wrap) {
  const c = cfg.colors;
  const s = cfg.fontSizes;
  return (
    "<tr>" +
    '<td width="44" style="width:44px;padding:0 8px 0 0;font-family:' + cfg.fontFamily + ";font-size:" + s.label + "px;line-height:" + s.rowLine +
    "px;font-weight:bold;color:" + c.label + ';white-space:nowrap;vertical-align:top;">' + escapeHtml(label) + "</td>" +
    // 地址允许换行：手机上屏幕窄，不换行会把左边的 Logo 挤扁
    '<td style="padding:0;font-family:' + cfg.fontFamily + ";font-size:" + s.row + "px;line-height:" + s.rowLine + "px;color:" + c.value +
    ";" + (wrap ? "" : "white-space:nowrap;") + 'vertical-align:top;">' + valueHtml + "</td>" +
    "</tr>"
  );
}

/** 一行免责声明（完整签名与精简签名共用同一段文字） */
function disclaimerHtml(cfg) {
  const co = cfg.company;
  if (!co.termsUrl && !co.disclaimerPrefix) return "";
  return (
    escapeHtml(co.disclaimerPrefix || "") + (co.disclaimerPrefix ? " " : "") +
    (co.termsUrl ? link(cfg, co.termsUrl, co.termsText || co.termsUrl, cfg.colors.footer, true) : "") + "."
  );
}

/**
 * 淡化圆环的背景图样式。背景图不占位、不会影响排版：
 * 支持的客户端把它画在文字下层，不支持的（经典 Outlook）或拦截远程图片的就什么都不显示。
 */
function markBackgroundCss(cfg, img, imageSrc, offset) {
  const o = offset || { right: 0, bottom: 0 };
  return (
    "background-image:url(" + escapeHtml(imageSrc("markCorner", true)) + ");background-repeat:no-repeat;" +
    "background-position:right " + o.right + "px bottom " + o.bottom + "px;" +
    "background-size:" + img.width + "px " + img.height + "px;"
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
  if (co.address) rows.push(row(cfg, L.address, co.addressUrl ? link(cfg, co.addressUrl, co.address, c.value) : plain(co.address), true));

  const s = cfg.fontSizes;
  const logo = cfg.images.logo;
  const badge = cfg.images.badge;
  const mark = cfg.images.mark;
  const corner = cfg.images.markCorner;
  const disclaimer = disclaimerHtml(cfg);
  const hasDisclaimer = !!disclaimer;

  let html = '<div id="acs-signature" data-acs-sig="full">';
  html += signOffHtml(cfg);
  html += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;">';
  html += "<tr>";

  // 左列：Logo，垂直居中
  if (logo) {
    html += '<td width="' + logo.width + '" valign="middle" style="padding:0 14px 0 0;vertical-align:middle;">' + imgTag(logo, imageSrc("logo")) + "</td>";
  }

  // 右列：姓名 / 职位 / 联系方式
  html += '<td valign="middle" style="border-left:2px solid ' + c.divider + ';padding:2px 0 2px 14px;vertical-align:middle;">';
  const nameHtml = '<div style="font-family:' + f + ";font-size:" + s.name + "px;line-height:" + (s.name + 4) + "px;font-weight:bold;color:" + c.text + ';">' + plain(p.displayName) + "</div>";
  if (mark && cfg.markPosition === "name") {
    // 品牌圆环放在姓名同一行的最右侧
    html += '<table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;"><tr>' +
      '<td valign="middle" style="vertical-align:middle;">' + nameHtml + "</td>" +
      '<td align="right" valign="middle" style="padding:0 0 0 16px;text-align:right;vertical-align:middle;">' + imgTag(mark, imageSrc("mark")) + "</td>" +
      "</tr></table>";
  } else {
    html += nameHtml;
  }
  if (cfg.showTitleInFull && p.jobTitle) {
    html += '<div style="font-family:' + f + ";font-size:" + s.title + "px;line-height:" + (s.title + 4) + "px;color:" + c.muted + ';padding:1px 0 0 0;">' +
      plain(p.jobTitle) + "</div>";
  }
  html += '<div style="height:7px;line-height:7px;font-size:1px;">&nbsp;</div>';
  const rowsHtml = '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;">' + rows.join("") + "</table>";
  if (corner && cfg.markPosition === "overlay") {
    // 淡化的 1/4 圆环做成 CSS 背景图，压在联系方式文字后面。
    // 用背景图而不是定位图片：手机版 Outlook 会删掉 position:absolute，图片就会掉进正文里破版；
    // 背景图最坏的情况只是不显示（经典 Outlook / 拦截远程图片时），排版永远不会乱。
    html += '<div style="' + markBackgroundCss(cfg, corner, imageSrc, cfg.markOverlayOffset) + '">' + rowsHtml + "</div>";
  } else if (corner && cfg.markPosition === "corner") {
    // 不重叠的版本：联系方式右侧单独一格（所有邮件客户端都支持）
    html += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;"><tr>' +
      '<td valign="bottom" style="vertical-align:bottom;">' + rowsHtml + "</td>" +
      '<td width="' + corner.width + '" align="right" valign="bottom" style="padding:0 0 2px 10px;text-align:right;vertical-align:bottom;">' +
      imgTag(corner, imageSrc("markCorner")) + "</td></tr></table>";
  } else {
    html += rowsHtml;
  }
  html += "</td></tr>";

  // 底部（跨两列，与上方同宽）：细分隔线 + WCA 徽章 + 一行免责声明 +（可选）品牌圆环
  const markInFooter = mark && cfg.markPosition === "footer";
  if (hasDisclaimer || badge || markInFooter) {
    html += '<tr><td colspan="' + (logo ? 2 : 1) + '" style="padding:11px 0 0 0;">';
    html += '<table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;"><tr>';
    html += '<td style="border-top:1px solid ' + c.hairline + ';padding:8px 0 0 0;">';
    html += '<table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;"><tr>';
    if (badge) {
      html += '<td width="' + badge.width + '" valign="middle" style="padding:0 12px 0 0;vertical-align:middle;">' + imgTag(badge, imageSrc("badge")) + "</td>";
    }
    if (hasDisclaimer) {
      html += '<td valign="middle" style="font-family:' + f + ";font-size:" + s.footer + "px;line-height:" + (s.footer + 4) + "px;color:" + c.footer +
        ';vertical-align:middle;">' + disclaimer + "</td>";
    }
    if (markInFooter) {
      html += '<td width="' + mark.width + '" align="right" valign="middle" style="padding:0 0 0 16px;text-align:right;vertical-align:middle;">' + imgTag(mark, imageSrc("mark")) + "</td>";
    }
    html += "</tr></table></td></tr></table></td></tr>";
  }
  html += "</table></div>";
  return html;
}

/**
 * 精简签名：三行 + 底部一行小字免责声明。用于同一会话中本人第二次及以后的回复。
 * markInShort=true 时，淡化的 1/4 圆环会压在网址那一行附近（经典 Outlook 自动隐藏）。
 * @param {(key:string)=>string} [imageSrc] 圆环图片地址；不传则不显示圆环
 */
export function buildShortSignature(p, cfg, imageSrc) {
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

  const corner = cfg.images.markCorner;
  const showMark = !!(imageSrc && corner && cfg.markInShort);
  const o = cfg.markShortOffset || { right: 0, bottom: 14 };

  let html = '<div id="acs-signature" data-acs-sig="short">';
  if (cfg.signOffInShort) html += signOffHtml(cfg);
  html += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;"><tr>';
  html += '<td style="border-left:3px solid ' + c.accent + ';padding:1px 0 1px 10px;font-family:' + f + ";font-size:12px;line-height:19px;color:" + c.muted + ';">';
  // 圆环用 CSS 背景图（不占位，手机上也不会把排版挤乱）
  if (showMark) html += '<div style="' + markBackgroundCss(cfg, corner, imageSrc, o) + '">';
  html += line1 + "<br/>" + line2 + "<br/>" + parts.join(sep);
  const disclaimer = disclaimerHtml(cfg);
  if (disclaimer) {
    html += '<div style="padding:6px 0 0 0;font-size:10px;line-height:14px;color:' + c.footer + ';">' + disclaimer + "</div>";
  }
  if (showMark) html += "</div>";
  html += "</td></tr></table></div>";
  return html;
}
