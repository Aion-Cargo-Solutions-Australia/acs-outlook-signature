/* global Office */
import { SIGNATURE_CONFIG as CFG } from "../config.js";
import { IMAGES } from "../generated/images.js";
import { ADDIN } from "../generated/addin.js";
import { buildFullSignature, buildShortSignature } from "./signature.js";
import { bodyHasOwnSignature, findSharedMailbox, normalizeProfile, teamProfile } from "./profile.js";
import { fetchMe, getToken, hasSentInConversation, signedInUser } from "./graph.js";
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
 * sessionData：只在这一封邮件的撰写过程中有效的小记事本。
 * 切换发件人时加载项会重新运行一次，靠它记住这封邮件里已经做过什么。
 */
async function sessionGet(key) {
  const it = item();
  if (!it.sessionData || !it.sessionData.getAsync) return "";
  try {
    return (await withTimeout(asPromise((cb) => it.sessionData.getAsync(key, cb)), 3000, "sessionData")) || "";
  } catch (e) {
    return ""; // 还没存过这个 key 时也会走到这里
  }
}

async function sessionSet(key, value) {
  const it = item();
  if (!it.sessionData || !it.sessionData.setAsync) return;
  try {
    await withTimeout(asPromise((cb) => it.sessionData.setAsync(key, value, cb)), 3000, "sessionData");
  } catch (e) {
    log("sessionData save failed", key, e && e.message);
  }
}

/**
 * 当前邮件的发件人 (From)。
 * @returns {{address:string, source:"from"|"mailbox"}} address 小写；source 说明地址是怎么拿到的：
 *   "from" = 读到了发件人栏，"mailbox" = 读不到发件人栏，退回当前邮箱的地址
 */
export async function getFrom() {
  const it = item();
  let from = "";
  if (it.from && it.from.getAsync) {
    try {
      const v = await withTimeout(asPromise((cb) => it.from.getAsync(cb)), 5000, "from");
      from = (v && v.emailAddress) || "";
    } catch (e) {
      log("from.getAsync failed", e && e.message);
    }
  }
  return {
    address: String(from || fallbackUser().emailAddress || "").toLowerCase(),
    source: from ? "from" : "mailbox",
  };
}

/**
 * 按发件人决定签名里用谁的资料：From 是 config.js 里登记的共享邮箱 → 团队版，否则是本人。
 * 结果里的 trace 是排查用的标记（代码版本 + 判断结果 + 当时看到的发件人），会写进签名容器的 id。
 * @param {object} profile loadProfile 得到的本人资料
 */
export async function resolveSender(profile) {
  const from = await getFrom();
  const box = findSharedMailbox(from.address, CFG);
  const sender = box ? teamProfile(profile, box, CFG) : Object.assign({}, profile);
  sender.trace = [ADDIN.build, box ? "team" : "personal", from.source, from.address].join(" ");
  return sender;
}

/**
 * 取得用户资料：优先 Graph（并写入缓存），失败则用缓存，最后用 Outlook 自带的姓名+邮箱。
 * @returns {{profile, token, error}}
 */
export async function loadProfile(allowPopup) {
  let cached = readCachedProfile();
  let token = null;
  let error = null;
  try {
    token = allowPopup ? await getToken(true) : await withTimeout(getToken(false), 8000, "token");
  } catch (e) {
    error = e;
    log("token failed", e && e.message);
  }
  // 共享邮箱在单独的窗口 / 账户里打开时，缓存可能存在共享邮箱里，是同事留下的 → 不是本人的缓存不用
  const user = signedInUser();
  if (cached && user && cached.profile.upn && cached.profile.upn !== user) cached = null;

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

const DECISIONS_KEY = "acsSigDecisions";

async function readDecisions() {
  try {
    return JSON.parse((await sessionGet(DECISIONS_KEY)) || "{}") || {};
  } catch (e) {
    return {};
  }
}

/**
 * 决定用完整签名还是精简签名。
 * @param {object} sender resolveSender 的结果（本人或团队版资料）
 * @param {{reapply?:boolean}} [opts] reapply=true：这封邮件里已经插入过签名（切换发件人 / 面板里重新插入）。
 *   这时正文里已经有加载项自己刚插入的签名，不能再拿正文判断，否则会把它当成“以前回复过”，所以沿用第一次的结论。
 * @returns {{variant:"full"|"short", reason:string, composeType:string}}
 */
export async function decideVariant(sender, token, opts) {
  const composeType = await getComposeType();
  if (composeType === "newMail") return { variant: "full", reason: "new message", composeType: composeType };
  if (composeType === "forward" && CFG.forwardAlwaysFull) return { variant: "full", reason: "forward", composeType: composeType };

  const reapply = !!(opts && opts.reapply);
  const decisions = reapply ? await readDecisions() : {};
  if (decisions[sender.email]) {
    return { variant: decisions[sender.email], reason: "decided earlier for this sender in this draft", composeType: composeType };
  }
  const decision = await decideReply(sender, token, composeType);
  decisions[sender.email] = decision.variant;
  await sessionSet(DECISIONS_KEY, JSON.stringify(decisions));
  return decision;
}

async function decideReply(sender, token, composeType) {
  // 共享邮箱：团队里谁回复过都算，所以不查“本人已发送邮件”，只看引用的历史邮件里有没有这个团队的签名
  const conversationId = item().conversationId;
  if (!sender.shared && token && conversationId) {
    try {
      if (await withTimeout(hasSentInConversation(token, conversationId), 6000, "sent-items")) {
        return { variant: "short", reason: "already sent in this conversation (Graph)", composeType: composeType };
      }
    } catch (e) {
      log("sent-items check failed", e && e.message);
    }
  }
  const body = await getBodyText();
  if (bodyHasOwnSignature(body, sender)) {
    return { variant: "short", reason: (sender.shared ? "team" : "own") + " signature found in quoted thread", composeType: composeType };
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

const EMBEDDED_KEY = "acsSigImages";

/** 这封邮件里已经由加载项内嵌、而且确实还在附件里的图片（切换发件人重新生成签名时不再重复附加） */
async function embeddedImages() {
  const it = item();
  if (!it.getAttachmentsAsync) return [];
  const saved = await sessionGet(EMBEDDED_KEY);
  if (!saved) return [];
  try {
    const list = (await withTimeout(asPromise((cb) => it.getAttachmentsAsync(cb)), 3000, "attachments")) || [];
    const present = list.filter((a) => a.isInline).map((a) => a.name);
    return saved.split(",").filter((name) => present.indexOf(name) !== -1);
  } catch (e) {
    return [];
  }
}

/**
 * 把签名写入当前撰写的邮件
 * @param {{reapply?:boolean}} [opts] reapply=true：这封邮件里已经插入过签名，现在是替换它
 */
export async function applySignature(variant, profile, opts) {
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
    const have = opts && opts.reapply ? await embeddedImages() : [];
    const added = [];
    // 只内嵌当前变体真正会用到的图片
    for (const k of imageKeysFor(variant)) {
      if (have.indexOf(IMAGES[k].cid) !== -1) continue;
      try {
        await addInlineImage(k);
        added.push(IMAGES[k].cid);
      } catch (e) {
        log("inline image failed, fallback to link", k, e && e.message);
        return applyHtml(buildHtml(variant, profile, "link"));
      }
    }
    if (added.length) await sessionSet(EMBEDDED_KEY, have.concat(added).join(","));
  }
  return applyHtml(buildHtml(variant, profile, mode));
}

function applyHtml(html) {
  return asPromise((cb) => item().body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, cb));
}
