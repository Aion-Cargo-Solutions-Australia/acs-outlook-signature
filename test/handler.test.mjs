import test from "node:test";
import assert from "node:assert/strict";
import { build } from "esbuild";
import vm from "node:vm";
import path from "node:path";

const root = path.resolve(import.meta.dirname, "..");
const bundle = (await build({
  entryPoints: [path.join(root, "src/launchevent/launchevent.js")], bundle: true, write: false, format: "iife", target: "es2017",
  plugins: [{ name: "mock", setup(b) { b.onResolve({ filter: /shared\/graph\.js$|\.\/graph\.js$/ }, () => ({ path: path.join(root, "test/mock-graph.js") })); } }],
})).outputFiles[0].text;

const IVY = { displayName: "Ivy Hu", mail: "ivy.hu@aioncargo.com", userPrincipalName: "ivy.hu@aioncargo.com", mobilePhone: "+61 425 666 802", businessPhones: ["+61 2 9160 2300 EXT 601"] };
const text = (html) => String(html || "").replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ");

/**
 * 模拟一封正在撰写的邮件。fire(name) 触发 manifest 里登记的事件；
 * 正文 = 引用的历史邮件 + 加载项当前插入的签名（和真实 Outlook 一样，签名插入后会出现在正文里）。
 */
function compose({ composeType, sent = false, tokenFail = false, body = "", from, mailbox, settings = {} }) {
  const calls = { attachments: [], signature: null, disabled: false, notices: 0 };
  const session = {};
  const handlers = {};
  const state = { from: from || "ivy.hu@aioncargo.com" };
  const ok = (value) => ({ status: "succeeded", value });
  const item = {
    conversationId: composeType === "newMail" ? null : "AAQkConv+/=",
    getComposeTypeAsync: (cb) => cb(ok({ composeType })),
    from: { getAsync: (cb) => cb(ok({ displayName: "", emailAddress: state.from })) },
    body: {
      getAsync: (t, cb) => cb(ok(body + "\n" + text(calls.signature))),
      setSignatureAsync: (html, o, cb) => { calls.signature = html; cb(ok()); },
    },
    sessionData: {
      getAsync: (k, cb) => cb(k in session ? ok(session[k]) : { status: "failed", error: new Error("not found") }),
      setAsync: (k, v, cb) => { session[k] = v; cb(ok()); },
    },
    getAttachmentsAsync: (cb) => cb(ok(calls.attachments.map((name) => ({ name, isInline: true })))),
    disableClientSignatureAsync: (cb) => { calls.disabled = true; cb(ok()); },
    addFileAttachmentFromBase64Async: (b64, name, o, cb) => { calls.attachments.push(name); cb(ok("id")); },
    notificationMessages: { addAsync: () => { calls.notices++; } },
  };
  const Office = {
    onReady: () => {}, AsyncResultStatus: { Failed: "failed", Succeeded: "succeeded" }, CoercionType: { Html: "html", Text: "text" },
    MailboxEnums: { ItemNotificationMessageType: { InsightMessage: "insightMessage" }, ActionType: { ShowTaskPane: "showTaskPane" } },
    actions: { associate: (n, f) => { handlers[n] = f; } },
    context: {
      requirements: { isSetSupported: () => true },
      roamingSettings: { get: (k) => settings[k], set: (k, v) => { settings[k] = v; }, saveAsync: (cb) => cb && cb(ok()) },
      mailbox: { item, userProfile: mailbox || { displayName: "Ivy Hu", emailAddress: "ivy.hu@aioncargo.com" } },
    },
  };
  const ctx = vm.createContext({ Office, console: { log() {} }, setTimeout, clearTimeout,
    __MOCK: { sent, tokenFail, me: IVY, user: IVY.userPrincipalName } });
  vm.runInContext(bundle, ctx);
  return {
    calls, mock: ctx.__MOCK, settings,
    setFrom: (address) => { state.from = address; },
    fire: (name) => new Promise((resolve) => handlers[name]({ completed: resolve })),
  };
}

async function run(options) {
  const c = compose(options);
  await c.fire("onNewMessageComposeHandler");
  return c;
}

test("new message → full signature with embedded images", async () => {
  const { calls } = await run({ composeType: "newMail" });
  assert.ok(calls.disabled);
  assert.match(calls.signature, /data-acs-sig="full"/);
  assert.deepEqual(calls.attachments, ["acs-logo.png", "wca-badge.png"]);
  assert.match(calls.signature, /cid:acs-logo\.png/);
});

test("first reply (never sent in conversation) → full", async () => {
  const { calls, mock } = await run({ composeType: "reply", sent: false });
  assert.match(calls.signature, /data-acs-sig="full"/);
  assert.equal(mock.checkedConv, "AAQkConv+/=");
});

test("later reply (already sent) → short, only the faded mark is embedded", async () => {
  const { calls } = await run({ composeType: "reply", sent: true });
  assert.match(calls.signature, /data-acs-sig="short"/);
  assert.equal(calls.attachments.length, 0); // 圆环是 CSS 背景图，不占附件
});

test("forward → full (forwardAlwaysFull)", async () => {
  const { calls } = await run({ composeType: "forward", sent: true });
  assert.match(calls.signature, /data-acs-sig="full"/);
});

test("no token → body fingerprint fallback + sign-in notice", async () => {
  const { calls } = await run({ composeType: "reply", tokenFail: true, body: "Kind regards,\nIvy Hu\nEmail ivy.hu@aioncargo.com" });
  assert.match(calls.signature, /data-acs-sig="short"/);
  assert.equal(calls.notices, 1);
});

test("shared mailbox: new message → team signature (team phone/email, no personal mobile)", async () => {
  const { calls } = await run({ composeType: "newMail", from: "AU.Importair@aioncargo.com" });
  assert.match(calls.signature, /data-acs-sig="full"/);
  for (const s of ["Ivy Hu", "ACS AU Import Air", "mailto:importair@aioncargo.com.au", "Ext.</span>&nbsp;821"]) assert.ok(calls.signature.includes(s), "missing " + s);
  for (const s of ["+61 425 666 802", "ivy.hu@aioncargo.com", "&nbsp;601"]) assert.ok(!calls.signature.includes(s), "should not contain " + s);
});

test("switching From to a shared mailbox replaces the signature without re-embedding images", async () => {
  const c = await run({ composeType: "newMail" });
  assert.ok(c.calls.signature.includes("+61 425 666 802"));
  c.setFrom("accounts@aioncargo.com.au");
  await c.fire("onMessageFromChangedHandler");
  assert.ok(c.calls.signature.includes("ACS AU Accounts") && !c.calls.signature.includes("+61 425 666 802"));
  assert.deepEqual(c.calls.attachments, ["acs-logo.png", "wca-badge.png"]);
  c.setFrom("ivy.hu@aioncargo.com");
  await c.fire("onMessageFromChangedHandler");
  assert.ok(c.calls.signature.includes("+61 425 666 802") && !c.calls.signature.includes("ACS AU Accounts"));
});

test("shared mailbox reply: full until the team's signature is in the thread, whoever sent it", async () => {
  const first = await run({ composeType: "reply", sent: true, from: "au.exportsea@aioncargo.com" });
  assert.match(first.calls.signature, /data-acs-sig="full"/);
  assert.equal(first.mock.checkedConv, undefined); // 不查本人的已发送邮件
  const later = await run({ composeType: "reply", from: "au.exportsea@aioncargo.com", body: "Kind regards,\nLuna Ye\nACS AU Export Sea\nEmail exportsea@aioncargo.com.au" });
  assert.match(later.calls.signature, /data-acs-sig="short"/);
  for (const s of ["Ivy Hu", "ACS AU Export Sea", "mailto:exportsea@aioncargo.com.au"]) assert.ok(later.calls.signature.includes(s), "missing " + s);
});

test("re-selecting the same From on a first reply keeps the full signature", async () => {
  const c = await run({ composeType: "reply", sent: false });
  assert.match(c.calls.signature, /data-acs-sig="full"/);
  await c.fire("onMessageFromChangedHandler");
  assert.match(c.calls.signature, /data-acs-sig="full"/);
  assert.equal(c.calls.attachments.length, 2);
});

test("shared mailbox opened on its own and no sign-in → team name instead of a person", async () => {
  const { calls } = await run({ composeType: "newMail", tokenFail: true, from: "au.accounts@aioncargo.com", mailbox: { displayName: "AU.Accounts", emailAddress: "au.accounts@aioncargo.com" } });
  assert.ok(calls.signature.includes(">ACS AU Accounts</div>"));
  assert.ok(!calls.signature.includes("AU.Accounts"));
});

test("a colleague's cached profile (shared mailbox settings) is not used for the signed-in user", async () => {
  const other = { displayName: "Someone Else", email: "someone.else@aioncargo.com", upn: "someone.else@aioncargo.com", mobile: "", phone: "+61 2 9160 2300", ext: "", source: "graph" };
  const settings = { acsSigProfile: JSON.stringify({ v: 3, savedAt: Date.now(), profile: other }) };
  const { calls } = await run({ composeType: "newMail", settings });
  assert.ok(calls.signature.includes("Ivy Hu") && !calls.signature.includes("Someone Else"));
});

test("the signature container id records build, decision and the From the add-in saw", async () => {
  const team = await run({ composeType: "newMail", from: "AU.Accounts@aioncargo.com" });
  assert.match(team.calls.signature, /<div id="acs-signature-[a-z0-9]+-team-from-au-accounts-aioncargo-com" data-acs-sig="full">/);
  const personal = await run({ composeType: "reply", sent: true });
  assert.match(personal.calls.signature, /<div id="acs-signature-[a-z0-9]+-personal-from-ivy-hu-aioncargo-com" data-acs-sig="short">/);
});
