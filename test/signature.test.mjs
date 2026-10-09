import test from "node:test";
import assert from "node:assert/strict";
import { SIGNATURE_CONFIG as CFG } from "../src/config.js";
import { buildFullSignature, buildShortSignature, escapeHtml } from "../src/shared/signature.js";
import { normalizeProfile, splitExtension, bodyHasOwnSignature, findSharedMailbox, teamProfile } from "../src/shared/profile.js";

const me = {
  displayName: "Ivy Hu", jobTitle: "Managing Director - Australia", mail: "Ivy.Hu@aioncargo.com",
  mobilePhone: "+61 425 666 802", businessPhones: ["+61 2 9160 2300 EXT 601"],
};
const src = (k) => "cid:" + k + ".png";

test("splitExtension handles common formats", () => {
  assert.deepEqual(splitExtension("+61 2 9160 2300 EXT 601"), { number: "+61 2 9160 2300", ext: "601" });
  assert.deepEqual(splitExtension("+61 2 9160 2300 x601"), { number: "+61 2 9160 2300", ext: "601" });
  assert.deepEqual(splitExtension("+61 2 9160 2300 ext.601"), { number: "+61 2 9160 2300", ext: "601" });
  assert.deepEqual(splitExtension("+61 2 9160 2300"), { number: "+61 2 9160 2300", ext: "" });
  assert.deepEqual(splitExtension(""), { number: "", ext: "" });
  assert.deepEqual(splitExtension("EXT 602"), { number: "", ext: "602" });
  assert.deepEqual(splitExtension("602"), { number: "", ext: "602" });
});

test("normalizeProfile: graph data, ext attribute and fallbacks", () => {
  const p = normalizeProfile(me, CFG, {});
  assert.equal(p.email, "ivy.hu@aioncargo.com");
  assert.equal(p.phone, "+61 2 9160 2300");
  assert.equal(p.ext, "601");
  assert.equal(p.source, "graph");
  const p2 = normalizeProfile({ displayName: "A B", onPremisesExtensionAttributes: { extensionAttribute1: "x 605" } }, CFG, {});
  assert.equal(p2.phone, CFG.company.mainPhone);
  assert.equal(p2.ext, "605");
  const p3 = normalizeProfile(null, CFG, { displayName: "Fallback User", emailAddress: "f@aioncargo.com" });
  assert.equal(p3.source, "fallback");
  assert.equal(p3.displayName, "Fallback User");
});

test("full signature contains all rows, images with explicit size, no <style>/svg", () => {
  const html = buildFullSignature(normalizeProfile(me, CFG, {}), CFG, src);
  for (const s of ["Ivy Hu", "+61 425 666 802", "Ext.</span>&nbsp;601", "mailto:ivy.hu@aioncargo.com", "www.aioncargo.com", "Pyrmont", "cid:logo.png", "cid:badge.png", "background-image:url(cid:markCorner.png)", 'width="150" height="91"', "Kind regards,", ">Mobile<", ">Tel<", ">Email<", ">Web<", ">Address<"]) {
    assert.ok(html.includes(s), "missing " + s);
  }
  assert.ok(!/<style|<svg|class=/i.test(html));
  assert.ok(html.length < 30000, "setSignatureAsync limit is 30,000 chars");
});

test("extension-only office phone uses main number; home street address is ignored", () => {
  const p = normalizeProfile({ displayName: "Leon Liu", businessPhones: ["EXT 602"], streetAddress: "11 Home St", city: "St Ives" }, CFG, {});
  assert.equal(p.phone, CFG.company.mainPhone);
  assert.equal(p.ext, "602");
  const html = buildFullSignature(p, CFG, src);
  assert.ok(!html.includes("Home St"));
  assert.ok(html.includes("Pyrmont"));
});

test("stale cached profile with a personal address never shows it", () => {
  const p = { ...normalizeProfile(me, CFG, {}), address: "11 Home St, St Ives NSW 2075" };
  for (const html of [buildFullSignature(p, CFG, src), buildShortSignature(p, CFG)]) {
    assert.ok(!html.includes("Home St"));
    assert.ok(html.includes("Pyrmont"));
  }
});

test("disclaimer is a single line", () => {
  const html = buildFullSignature(normalizeProfile(me, CFG, {}), CFG, src);
  const m = html.match(/This email is subject to our.*?<\/td>/);
  assert.ok(m && !/<br/i.test(m[0]));
});

test("full signature omits empty rows", () => {
  const html = buildFullSignature(normalizeProfile({ displayName: "No Mobile", mail: "n@aioncargo.com" }, CFG, {}), CFG, src);
  assert.ok(!html.includes(">Mobile<"));
  assert.ok(html.includes(">Tel<") && html.includes("+61 2 9160 2300")); // falls back to main phone
});

test("short signature: no <img>, faded mark only as CSS background", () => {
  const html = buildShortSignature(normalizeProfile(me, CFG, {}), CFG, src);
  assert.ok(!/<img/i.test(html));
  assert.ok(html.includes("background-image:url(") && html.includes("background-repeat:no-repeat"));
  assert.ok(!/background-image/i.test(buildShortSignature(normalizeProfile(me, CFG, {}), CFG))); // 不传 imageSrc → 纯文字
  for (const s of ["Ivy Hu", "Managing Director - Australia", "Mobile</b>&nbsp;+61 425 666 802", "Ext.</span>&nbsp;601", "Aion Cargo Solutions", "Suite 4.06, 55 Miller St, Pyrmont NSW 2009"]) assert.ok(html.includes(s), "missing " + s);
});

test("short signature carries the same one-line disclaimer, no images", () => {
  const html = buildShortSignature(normalizeProfile(me, CFG, {}), CFG, src);
  assert.ok(html.includes(CFG.company.disclaimerPrefix));
  assert.ok(html.includes(CFG.company.termsUrl) && html.includes(escapeHtml(CFG.company.termsText)));
  assert.ok(!/<br\s*\/?>\s*$/i.test(html));
});

test("the faded mark never takes part in layout (background image, no positioning)", () => {
  const html = buildFullSignature(normalizeProfile(me, CFG, {}), CFG, src);
  assert.ok(!/position\s*:\s*absolute/i.test(html));
  assert.match(html, /background-position:right \d+px bottom -?\d+px/);
});

test("HTML injection is escaped", () => {
  assert.equal(escapeHtml('<b>"x"&'), "&lt;b&gt;&quot;x&quot;&amp;");
  const html = buildShortSignature(normalizeProfile({ displayName: "<script>alert(1)</script>" }, CFG, {}), CFG);
  assert.ok(!html.includes("<script>"));
});

test("thread detection: own signature vs quoted headers", () => {
  const p = normalizeProfile(me, CFG, {});
  const firstReply = "Hi Ivy,\n\nFrom: John <john@client.com>\nTo: Ivy Hu <ivy.hu@aioncargo.com>\nSubject: quote\n\nPlease quote.";
  assert.equal(bodyHasOwnSignature(firstReply, p), false);
  const laterReply = firstReply + "\n\nKind regards,\nIvy Hu\nMob\t+61 425 666 802\nEmail\tivy.hu@aioncargo.com";
  assert.equal(bodyHasOwnSignature(laterReply, p), true);
  const localFormat = "Ivy Hu | M 0425 666 802";
  assert.equal(bodyHasOwnSignature(localFormat, p), true);
  const emailOnly = normalizeProfile({ displayName: "X", mail: "x@aioncargo.com" }, CFG, {});
  assert.equal(bodyHasOwnSignature("Email x@aioncargo.com", emailOnly), true);
  assert.equal(bodyHasOwnSignature("To: x@aioncargo.com", emailOnly), false);
});

test("shared mailbox table: every address belongs to exactly one team", () => {
  const seen = new Map();
  for (const box of CFG.sharedMailboxes) {
    assert.ok(box.team && box.phone && box.addresses.length, "incomplete entry " + JSON.stringify(box));
    for (const a of box.addresses) {
      const key = a.toLowerCase();
      assert.ok(!seen.has(key), key + " is listed for both " + seen.get(key) + " and " + box.team);
      seen.set(key, box.team);
    }
  }
});

test("findSharedMailbox matches any listed address, case-insensitively", () => {
  assert.equal(findSharedMailbox("importair@aioncargo.com.au", CFG).team, "ACS AU Import Air");
  assert.equal(findSharedMailbox(" AU.ImportSea@AionCargo.com ", CFG).team, "ACS AU Import Sea");
  assert.equal(findSharedMailbox("ivy.hu@aioncargo.com", CFG), null);
  assert.equal(findSharedMailbox("", CFG), null);
});

test("teamProfile keeps the person's name, swaps in team contact details", () => {
  const box = findSharedMailbox("au.accounts@aioncargo.com", CFG);
  const t = teamProfile(normalizeProfile(me, CFG, {}), box, CFG);
  assert.deepEqual([t.displayName, t.team, t.email, t.phone, t.ext, t.mobile, t.shared], ["Ivy Hu", "ACS AU Accounts", "accounts@aioncargo.com.au", "+61 2 9160 2300", "820", "", true]);
  const withMobile = teamProfile(normalizeProfile(me, CFG, {}), { ...box, showMobile: true }, CFG);
  assert.equal(withMobile.mobile, "+61 425 666 802");
});

test("team signatures: full and short show team name and team mailbox", () => {
  const t = teamProfile(normalizeProfile(me, CFG, {}), findSharedMailbox("exportair@aioncargo.com.au", CFG), CFG);
  const full = buildFullSignature(t, CFG, src);
  for (const s of ["Ivy Hu", "Managing Director - Australia", "ACS AU Export Air", "mailto:exportair@aioncargo.com.au", "&nbsp;822", "Pyrmont"]) assert.ok(full.includes(s), "full missing " + s);
  assert.ok(!full.includes(">Mobile<") && full.length < 30000);
  const short = buildShortSignature(t, CFG);
  for (const s of ["Ivy Hu", "ACS AU Export Air", "mailto:exportair@aioncargo.com.au", "&nbsp;822"]) assert.ok(short.includes(s), "short missing " + s);
  // 个人签名不受影响：精简版不带邮箱
  assert.ok(!buildShortSignature(normalizeProfile(me, CFG, {}), CFG).includes("mailto:"));
});

test("thread detection: team signature is recognised by the team mailbox only", () => {
  const t = teamProfile(normalizeProfile(me, CFG, {}), { ...findSharedMailbox("accounts@aioncargo.com.au", CFG), showMobile: true }, CFG);
  assert.equal(bodyHasOwnSignature("Kind regards,\nZora Zhang\nACS AU Accounts\nEmail accounts@aioncargo.com.au", t), true);
  assert.equal(bodyHasOwnSignature("Kind regards,\nIvy Hu\nMobile +61 425 666 802\nEmail ivy.hu@aioncargo.com", t), false);
});
