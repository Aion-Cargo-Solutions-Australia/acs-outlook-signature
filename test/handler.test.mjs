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

function run({ composeType, sent = false, tokenFail = false, body = "" }) {
  return new Promise((resolve) => {
    const calls = { attachments: [], signature: null, disabled: false, notices: 0 };
    const settings = {};
    let handler;
    const ok = (value) => ({ status: "succeeded", value });
    const item = {
      conversationId: composeType === "newMail" ? null : "AAQkConv+/=",
      getComposeTypeAsync: (cb) => cb(ok({ composeType })),
      body: {
        getAsync: (t, cb) => cb(ok(body)),
        setSignatureAsync: (html, o, cb) => { calls.signature = html; cb(ok()); },
      },
      disableClientSignatureAsync: (cb) => { calls.disabled = true; cb(ok()); },
      addFileAttachmentFromBase64Async: (b64, name, o, cb) => { calls.attachments.push(name); cb(ok("id")); },
      notificationMessages: { addAsync: () => { calls.notices++; } },
    };
    const Office = {
      onReady: () => {}, AsyncResultStatus: { Failed: "failed", Succeeded: "succeeded" }, CoercionType: { Html: "html", Text: "text" },
      MailboxEnums: { ItemNotificationMessageType: { InsightMessage: "insightMessage" }, ActionType: { ShowTaskPane: "showTaskPane" } },
      actions: { associate: (n, f) => { handler = f; } },
      context: {
        requirements: { isSetSupported: () => true },
        roamingSettings: { get: (k) => settings[k], set: (k, v) => { settings[k] = v; }, saveAsync: (cb) => cb && cb(ok()) },
        mailbox: { item, userProfile: { displayName: "Ivy Hu", emailAddress: "ivy.hu@aioncargo.com" } },
      },
    };
    const ctx = vm.createContext({ Office, console: { log() {} }, setTimeout, clearTimeout,
      __MOCK: { sent, tokenFail, me: { displayName: "Ivy Hu", mail: "ivy.hu@aioncargo.com", mobilePhone: "+61 425 666 802", businessPhones: ["+61 2 9160 2300 EXT 601"] } } });
    vm.runInContext(bundle, ctx);
    handler({ completed: () => resolve({ calls, mock: ctx.__MOCK }) });
  });
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

test("later reply (already sent) → short, no images", async () => {
  const { calls } = await run({ composeType: "reply", sent: true });
  assert.match(calls.signature, /data-acs-sig="short"/);
  assert.equal(calls.attachments.length, 0);
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
