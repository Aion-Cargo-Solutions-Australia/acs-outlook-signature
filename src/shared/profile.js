/**
 * 用户资料：Graph /me → 标准化对象。纯函数部分可在 Node 中测试。
 */

export const GRAPH_SELECT = [
  "displayName", "givenName", "surname", "jobTitle", "department", "mail", "userPrincipalName",
  "mobilePhone", "businessPhones", "officeLocation", "streetAddress", "city", "state", "postalCode",
  "onPremisesExtensionAttributes",
].join(",");

/** "+61 2 9160 2300 EXT 601" / "+61 2 9160 2300 x601" / "+61 2 9160 2300 ext.601" → { number, ext } */
export function splitExtension(raw) {
  if (!raw) return { number: "", ext: "" };
  const m = String(raw).match(/^(.*?)[\s,;]*(?:ext\.?|extension|x|#)\s*(\d{1,6})\s*$/i);
  if (m && m[1].replace(/\D/g, "").length >= 6) return { number: m[1].trim(), ext: m[2] };
  return { number: String(raw).trim(), ext: "" };
}

/**
 * @param {object} me  Graph /me 返回（可能为 null）
 * @param {object} cfg SIGNATURE_CONFIG
 * @param {object} fallback { displayName, emailAddress } —— 来自 Office.context.mailbox.userProfile
 */
export function normalizeProfile(me, cfg, fallback) {
  me = me || {};
  fallback = fallback || {};
  const bp = (me.businessPhones && me.businessPhones[0]) || "";
  const split = splitExtension(bp);
  let ext = split.ext;
  if (!ext && me.onPremisesExtensionAttributes && cfg.extensionAttribute) {
    const v = me.onPremisesExtensionAttributes[cfg.extensionAttribute];
    if (v) ext = String(v).replace(/^\s*(ext\.?|x)\s*/i, "").trim();
  }
  return {
    displayName: me.displayName || fallback.displayName || "",
    jobTitle: me.jobTitle || "",
    department: me.department || "",
    email: (me.mail || fallback.emailAddress || me.userPrincipalName || "").toLowerCase(),
    mobile: (me.mobilePhone || "").trim(),
    phone: split.number || (cfg.company && cfg.company.mainPhone) || "",
    ext: ext || "",
    // 只有在 Entra 填了街道地址时才覆盖公司地址
    address: me.streetAddress
      ? [me.streetAddress, [me.city, me.state, me.postalCode].filter(Boolean).join(" ")].filter(Boolean).join(", ")
      : "",
    source: me.displayName ? "graph" : "fallback",
  };
}

function digits(s) {
  return String(s || "").replace(/\D/g, "");
}

/**
 * 判断引用的历史邮件正文里是否已经出现过“本人”的签名。
 * 用于 Graph 不可用时的兜底判断。
 * 指纹：本人手机号（纯数字）或 “Email <本人邮箱>” 这一行（邮件头的 From/To 行不会以 "Email" 开头）。
 */
export function bodyHasOwnSignature(bodyText, profile) {
  if (!bodyText || !profile) return false;
  const text = String(bodyText);
  const mob = digits(profile.mobile);
  if (mob.length >= 8) {
    const bodyDigits = digits(text);
    // 同时匹配国际格式和本地格式（+61 4xx / 04xx）
    const local = mob.replace(/^61/, "0");
    if (bodyDigits.indexOf(mob) !== -1 || (local !== mob && bodyDigits.indexOf(local) !== -1)) return true;
  }
  if (profile.email) {
    const e = profile.email.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    if (new RegExp("(^|\\s)Email\\s*[:：]?\\s*<?(mailto:)?" + e, "i").test(text)) return true;
  }
  return false;
}
