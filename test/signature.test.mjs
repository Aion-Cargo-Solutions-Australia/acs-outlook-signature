import test from "node:test";
import assert from "node:assert/strict";
import { SIGNATURE_CONFIG as CFG } from "../src/config.js";
import { activeBanner, buildFullSignature, buildShortSignature, escapeHtml } from "../src/shared/signature.js";
import { normalizeProfile, splitExtension, bodyHasOwnSignature, findSharedMailbox, teamProfile, configForOffice } from "../src/shared/profile.js";

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
  const m = html.match(/including any ACS AU quotation.*?<\/td>/);
  assert.ok(m && !/<br/i.test(m[0]));
});

test("full signature uses no colspan: the wide disclaimer below must not stretch the logo column in Word", () => {
  const sh = normalizeProfile({ displayName: "A B", officeLocation: "Shanghai" }, CFG, {});
  for (const [p, cfg] of [[normalizeProfile(me, CFG, {}), CFG], [sh, configForOffice(sh.office, CFG)]]) {
    assert.ok(!buildFullSignature(p, cfg, src).includes("colspan"));
  }
});

test("full signature omits empty rows", () => {
  const html = buildFullSignature(normalizeProfile({ displayName: "No Mobile", mail: "n@aioncargo.com" }, CFG, {}), CFG, src);
  assert.ok(!html.includes(">Mobile<"));
  assert.ok(html.includes(">Tel<") && html.includes("+61 2 9160 2300")); // falls back to main phone
});

test("short signature: no <img> except the seasonal banner, faded mark only as CSS background", () => {
  const html = buildShortSignature(normalizeProfile(me, CFG, {}), CFG, src);
  assert.equal((html.match(/<img/gi) || []).length, activeBanner(CFG) ? 1 : 0);
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

const LEON = {
  displayName: "闾震涛 Leon", jobTitle: "Branch Manager", mail: "Leon.lv@aioncargo.com", officeLocation: "Shanghai",
  mobilePhone: "+86 18616172470", businessPhones: [],
};

test("configForOffice: Entra office picks the office settings, anything else keeps the defaults", () => {
  for (const office of ["Shanghai", "SHANGHAI 813", "上海"]) assert.equal(configForOffice(office, CFG).company.mainPhone, "+86 21 65050600", office);
  for (const office of ["Sydney", "", null, undefined]) assert.equal(configForOffice(office, CFG), CFG);
  // 只覆盖写了的字段：标签、网址、Logo 沿用默认
  const sh = configForOffice("Shanghai", CFG);
  assert.equal(sh.labels.mobile, CFG.labels.mobile);
  assert.equal(sh.company.website, CFG.company.website);
  assert.equal(sh.images.logo, CFG.images.logo);
  assert.equal(sh.images.badge, null); // 上海暂不放 WCA 徽章
  assert.ok(CFG.images.badge);
});

test("Shanghai signature: Chinese entity, Shanghai address and phone, its own T&C; no Australian details", () => {
  const p = normalizeProfile(LEON, CFG, {});
  assert.equal(p.phone, "+86 21 65050600");
  const cfg = configForOffice(p.office, CFG);
  const full = buildFullSignature(p, cfg, src);
  const short = buildShortSignature(p, cfg, src);
  for (const s of ["闾震涛 Leon", "Branch Manager", "艾朗国际物流（上海）有限公司", "上海市虹口区衡水路151号前海人寿大厦813室", "+86 18616172470",
    "+86 21 65050600", ">Website<", "ACS-Terms-and-Conditions.pdf", "is subject to ACS <a", "Microsoft YaHei", "cid:logo.png"]) assert.ok(full.includes(s), "full missing " + s);
  assert.ok(!full.includes("cid:badge.png"));
  for (const s of ["艾朗国际物流（上海）有限公司", "上海市虹口区衡水路151号前海人寿大厦813室", "ACS-Terms-and-Conditions.pdf"]) assert.ok(short.includes(s), "short missing " + s);
  for (const s of ["Pyrmont", "+61", "au-tc", "maps.google", "ACS AU"]) assert.ok(!full.includes(s) && !short.includes(s), "should not contain " + s);
  // 澳洲签名不显示公司名这一行
  assert.ok(!buildFullSignature(normalizeProfile(me, CFG, {}), CFG, src).includes(">Aion Cargo Solutions</div>"));
});

test("personal notes: shown in red before the name for that person only, in full and short", () => {
  const candy = normalizeProfile({ displayName: "Candy Wang", mail: "Candy.Wang@aioncargo.com", officeLocation: "Shanghai" }, CFG, {});
  const cfg = configForOffice(candy.office, CFG);
  for (const html of [buildFullSignature(candy, cfg, src), buildShortSignature(candy, cfg, src)]) {
    assert.ok(html.includes("一次性补料，任何更改都会收取改单费RMB500"));
    assert.ok(html.includes("<b>马士基VGM仅限一次性申报，更新会有罚金USD300</b>"));
    assert.ok(html.indexOf("改单费") < html.indexOf("Candy Wang"), "notes go above the name");
    assert.ok(html.includes("color:" + CFG.colors.note));
  }
  const others = [normalizeProfile(me, CFG, {}), teamProfile(candy, findSharedMailbox(CFG.sharedMailboxes[0].addresses[0], CFG), CFG)];
  for (const p of others) assert.ok(!buildFullSignature(p, CFG, src).includes("改单费"));
});

test("New Zealand office: Auckland address and phone, DDI label, ACS T&C; no Australian details", () => {
  const ivan = normalizeProfile({ displayName: "Ivan Qiu", jobTitle: "Sales & BD Director", mail: "ivan.qiu@aioncargo.com", officeLocation: "Auckland", mobilePhone: "+64 22 040 7673", businessPhones: ["+64 9 254 4907"] }, CFG, {});
  const cfg = configForOffice(ivan.office, CFG);
  assert.equal(configForOffice("New Zealand", CFG).company.mainPhone, "+64 9 254 4900");
  const full = buildFullSignature(ivan, cfg, src);
  for (const s of ["Ivan Qiu", "Sales &amp; BD Director", ">DDI<", "+64 9 254 4907", "40 Airpark Drive, Mangere, Auckland 2022", "ACS-Terms-and-Conditions.pdf", "cid:iata.png", "cid:cbaff.png", "cid:wcaNz.png"]) assert.ok(full.includes(s), "missing " + s);
  assert.ok(full.indexOf("The content of this email") < full.indexOf("cid:iata.png"), "badges go below the disclaimer");
  for (const s of ["Pyrmont", "+61", "au-tc", "ACS AU", "cid:badge.png"]) assert.ok(!full.includes(s), "should not contain " + s);
  assert.equal(normalizeProfile({ displayName: "A B", officeLocation: "Auckland" }, CFG, {}).phone, "+64 9 254 4900");
});

test("footer badges: one sits left of the disclaimer, several go on their own row below it", () => {
  const p = normalizeProfile(me, CFG, {});
  const one = buildFullSignature(p, CFG, src);
  assert.ok(one.indexOf("cid:badge.png") < one.indexOf("The content of this email"));
  const cfg = Object.assign({}, CFG, { images: Object.assign({}, CFG.images, { iata: CFG.images.badge }), footerBadges: ["iata", "badge"] });
  const two = buildFullSignature(p, cfg, src);
  assert.ok(two.indexOf("The content of this email") < two.indexOf("cid:iata.png"));
  assert.ok(two.indexOf("cid:iata.png") < two.indexOf("cid:badge.png"));
});

test("Christmas banner: AU personal and team signatures (full + short) until it expires, never other offices", () => {
  const withUntil = (until) => ({ ...CFG, images: { ...CFG.images, banner: { ...CFG.images.banner, until } } });
  const live = withUntil("2999-01-01T00:00:00Z");
  const p = normalizeProfile(me, live, {});
  const team = teamProfile(p, findSharedMailbox("accounts@aioncargo.com.au", live), live);
  for (const html of [buildFullSignature(p, live, src), buildShortSignature(p, live, src), buildFullSignature(team, live, src), buildShortSignature(team, live, src)]) {
    assert.ok(html.includes('src="cid:banner.png" width="480" height="88"'));
  }
  const gone = withUntil("2000-01-01T00:00:00Z");
  assert.ok(!buildFullSignature(p, gone, src).includes("banner") && !buildShortSignature(p, gone, src).includes("banner"));
  for (const office of ["Auckland", "Shanghai"]) assert.equal(activeBanner(configForOffice(office, live)), null, office);
  // 1 月 1 日（悉尼时间）当天还显示，1 月 2 日 0 点起不再显示
  assert.equal(new Date(CFG.images.banner.until).toISOString(), "2027-01-01T13:00:00.000Z");
});
