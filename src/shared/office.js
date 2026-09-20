/* global Office */
import { SIGNATURE_CONFIG as CFG } from "../config.js";
import { IMAGES } from "../generated/images.js";
import { ADDIN } from "../generated/addin.js";
import { buildFullSignature, buildShortSignature } from "./signature.js";
import { bodyHasOwnSignature, normalizeProfile } from "./profile.js";
import { fetchMe, getToken, hasSentInConversation } from "./graph.js";
import { readCachedProfile, isFresh, saveCachedProfile } from "./cache.js";

function asPromise(fn) {
  return new Promise((resolve, reject) => {
    fn((r) => {
      if (r.status === Office.AsyncResultStatus.Failed) reject(r.error);
      else resolve(r.value);
    });
  });
}

/** 给异步操作加超时，避免登录 / Graph 卡住导致整封邮件没有签名 */
export function withTimeout(promise, ms, label) {
  return new Promise((resolve, reject) => {
    const t = setTimeout(() => reject(new Error((label || "operation") + " timed out after " + ms + "ms")), ms);
    promise.then(
      (v) => { clearTimeout(t); resolve(v); },
      (e) => { clearTimeout(t); reject(e); }
    );
  });
}

function item() {
  return Office.context.mailbox.item;
}

export function log() {
  // 事件运行时里的 console.log 会写入 runtime log，便于排查
  try {
    console.log.apply(console, ["[ACS-SIG]"].concat(Array.prototype.slice.call(arguments)));
  } catch (e) {
    /* ignore */
  }
}

export async function getComposeType() {
  const it = item();
  if (!it.getComposeTypeAsync) return "newMail";
  try {
    const v = await asPromise((cb) => it.getComposeTypeAsync(cb));
    return (v && v.composeType) || "newMail";
  } catch (e) {
    return "newMail";
  }
}

export async function getBodyText() {
  try {
    return await asPromise((cb) => item().body.getAsync(Office.CoercionType.Text, cb));
  } catch (e) {
    return "";
  }
}

function fallbackUser() {
  const up = Office.context.mailbox.userProfile || {};
  return { displayName: up.displayName, emailAddress: up.emailAddress };
}

/**
 * 取得用户资料：优先 Graph（并写入缓存），失败则用缓存，最后用 Outlook 自带的姓名+邮箱。
 * @returns {{profile, token, error}}
 */
export async function loadProfile(allowPopup) {
  const cached = readCachedProfile();
  let token = null;
  let error = null;
  try {
    token = allowPopup ? await getToken(true) : await withTimeout(getToken(false), 8000, "token");
  } catch (e) {
    error = e;
    log("token failed", e && e.message);
  }

  if (token && (!isFresh(cached) || allowPopup)) {
    try {
      const me = await withTimeout(fetchMe(token), 8000, "graph /me");
      const profile = normalizeProfile(me, CFG, fallbackUser());
      await saveCachedProfile(profile);
      return { profile: profile, token: token, error: null };
    } catch (e) {
      error = e;
      log("graph /me failed", e && e.message);
    }
  }
  if (cached) return { profile: cached.profile, token: token, error: error };
  return { profile: normalizeProfile(null, CFG, fallbackUser()), token: token, error: error };
}

/**
 * 决定用完整签名还是精简签名。
 * @returns {{variant:"full"|"short", reason:string, composeType:string}}
 */
export async function decideVariant(profile, token) {
  const composeType = await getComposeType();
  if (composeType === "newMail") return { variant: "full", reason: "new message", composeType: composeType };
  if (composeType === "forward" && CFG.forwardAlwaysFull) return { variant: "full", reason: "forward", composeType: composeType };

  const conversationId = item().conversationId;
  if (token && conversationId) {
    try {
      if (await withTimeout(hasSentInConversation(token, conversationId), 6000, "sent-items")) {
        return { variant: "short", reason: "already sent in this conversation (Graph)", composeType: composeType };
      }
    } catch (e) {
      log("sent-items check failed", e && e.message);
    }
  }
  const body = await getBodyText();
  if (bodyHasOwnSignature(body, profile)) {
    return { variant: "short", reason: "own signature found in quoted thread", composeType: composeType };
  }
  return { variant: "full", reason: "first reply in this conversation", composeType: composeType };
}

async function addInlineImage(key) {
  const img = IMAGES[key];
  if (!img) return;
  await asPromise((cb) => item().addFileAttachmentFromBase64Async(img.base64, img.cid, { isInline: true }, cb));
}

export function imageSrcFactory(mode) {
  // forceLink=true：CSS 背景图只能用 http(s) 地址，cid: 在背景图里不生效
  return function (key, forceLink) {
    const img = IMAGES[key];
    if (!img) return "";
    return mode === "embed" && !forceLink ? "cid:" + img.cid : ADDIN.baseUrl + "/assets/" + img.file;
  };
}

/** 当前变体真正需要内嵌（CID 附件）的图片。圆环走 CSS 背景图，不需要内嵌 */
export function imageKeysFor(variant) {
  const pos = CFG.markPosition;
  if (variant === "short") return [];
  const used = (k) => (k === "mark" ? pos === "footer" || pos === "name" : k === "markCorner" ? pos === "corner" : true);
  return Object.keys(IMAGES).filter((k) => CFG.images[k] && used(k));
}

export function buildHtml(variant, profile, mode) {
  const src = imageSrcFactory(mode || CFG.imageMode);
  return variant === "short"
    ? buildShortSignature(profile, CFG, CFG.markInShort ? src : null)
    : buildFullSignature(profile, CFG, src);
}

/** 把签名写入当前撰写的邮件 */
export async function applySignature(variant, profile) {
  const it = item();
  // 关闭 Outlook 客户端自己保存的签名，避免出现两份签名
  if (it.disableClientSignatureAsync) {
    try {
      await asPromise((cb) => it.disableClientSignatureAsync(cb));
    } catch (e) {
      log("disableClientSignature failed", e && e.message);
    }
  }
  const mode = CFG.imageMode;
  if (mode === "embed") {
    // 只内嵌当前变体真正会用到的图片
    for (const k of imageKeysFor(variant)) {
      try {
        await addInlineImage(k);
      } catch (e) {
        log("inline image failed, fallback to link", k, e && e.message);
        return applyHtml(buildHtml(variant, profile, "link"));
      }
    }
  }
  return applyHtml(buildHtml(variant, profile, mode));
}

function applyHtml(html) {
  return asPromise((cb) => item().body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, cb));
}
