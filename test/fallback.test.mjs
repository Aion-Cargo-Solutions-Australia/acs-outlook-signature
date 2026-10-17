import test from "node:test";
import assert from "node:assert/strict";
import { SIGNATURE_CONFIG as CFG } from "../src/config.js";
import { activeBanner, buildFullSignature, buildShortSignature } from "../src/shared/signature.js";
import { normalizeProfile, teamProfile } from "../src/shared/profile.js";
import { buildFallbackHtml, buildHolidayHtml, holidayPattern, markerLine, markerPattern, mobileSignatureText, renderMobilePage, renderRulesScript } from "../src/shared/fallback.js";

// 近似 Exchange 匹配正文时看到的纯文本：去掉标签、解码常见实体、合并空白
const text = (html) =>
  String(html).replace(/<br\s*\/?>/gi, "\n").replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">");
const exchangeMatches = (pattern, html) => new RegExp(pattern, "i").test(text(html));

const me = { displayName: "Ivy Hu", jobTitle: "Managing Director", mail: "ivy.hu@aioncargo.com", mobilePhone: "+61 425 666 802", businessPhones: ["+61 2 9160 2300 EXT 601"] };

for (const box of CFG.sharedMailboxes) {
  test("fallback for " + box.team + ": block, mobile text and rule pattern agree", () => {
    const pattern = markerPattern(box, CFG);
    const html = buildFallbackHtml(box, CFG);
    assert.ok(html.length < 5000, "Exchange disclaimer limit is 5000 characters");
    assert.ok(!/<img|<style|class=/i.test(html));
    for (const s of [box.team, box.addresses[0], "Ext. " + box.ext, CFG.company.website]) assert.ok(text(html).includes(s), "block missing " + s);
    // 兜底内容和手机签名都会被规则识别为“已有签名”，同一封邮件 / 同一个 chain 里不会再追加
    assert.ok(exchangeMatches(pattern, html), "rule must recognise its own block");
    assert.ok(exchangeMatches(pattern, mobileSignatureText(box, CFG).replace(/\n/g, "<br>")), "rule must recognise the mobile signature");
    assert.ok(mobileSignatureText(box, CFG).split("\n").includes(markerLine(box, CFG)));
  });

  test("fallback for " + box.team + ": the add-in's own signatures and quoted headers don't trip the rule", () => {
    const pattern = markerPattern(box, CFG);
    const team = teamProfile(normalizeProfile(me, CFG, {}), box, CFG);
    const src = (k) => "cid:" + k + ".png";
    assert.ok(!exchangeMatches(pattern, buildFullSignature(team, CFG, src)));
    assert.ok(!exchangeMatches(pattern, buildShortSignature(team, CFG, src)));
    // 回复 / 转发里引用的邮件头
    assert.ok(!exchangeMatches(pattern, "From: John &lt;john@client.com&gt;<br>To: AU.Accounts &lt;" + box.addresses[0] + "&gt;<br>Subject: invoice"));
  });
}

test("rules script: one idempotent rule per shared mailbox with both exceptions, ASCII only", () => {
  const ps = renderRulesScript(CFG);
  assert.ok(/^[\x00-\x7f]*$/.test(ps), "Windows PowerShell 5.1 reads scripts without BOM as ANSI");
  const holiday = activeBanner(CFG) ? 1 : 0;
  assert.equal((ps.match(/^ {4}Name {4}= "ACS signature fallback - /gm) || []).length, CFG.sharedMailboxes.length + holiday);
  for (const box of CFG.sharedMailboxes) {
    assert.ok(ps.includes('From    = @("' + box.addresses[0] + '")'));
    assert.ok(ps.includes("Pattern = '" + markerPattern(box, CFG) + "'"));
  }
  for (const s of ['ApplyHtmlDisclaimerLocation          = "Append"', 'ApplyHtmlDisclaimerFallbackAction    = "Ignore"',
    'ExceptIfHeaderContainsMessageHeader  = "X-ACS-Signature"', 'ExceptIfHeaderContainsWords          = "inserted"',
    "Set-TransportRule -Identity $r.Name @params", "New-TransportRule -Name $r.Name @params"]) {
    assert.ok(ps.includes(s), "missing " + s);
  }
  // here-string 的结束符必须顶格，否则 PowerShell 解析失败
  assert.equal((ps.match(/^'@$/gm) || []).length, CFG.sharedMailboxes.length + holiday);
  // 单引号字符串里不能出现单引号
  for (const box of CFG.sharedMailboxes) assert.ok(!markerPattern(box, CFG).includes("'"));
});

test("mobile page lists every team with a copyable signature", () => {
  const page = renderMobilePage(CFG);
  for (const box of CFG.sharedMailboxes) {
    assert.ok(page.includes(box.addresses[0]));
    assert.ok(page.includes(markerLine(box, CFG).replace("&", "&amp;")));
  }
  assert.ok(page.includes("[Your name]"));
});

test("holiday notice rule: text version of the banner for every shared mailbox, expires with it", () => {
  const withUntil = (until) => ({ ...CFG, images: { ...CFG.images, banner: { ...CFG.images.banner, until } } });
  const live = withUntil("2027-01-02T00:00:00+11:00");
  const ps = renderRulesScript(live);
  assert.ok(ps.includes('Name    = "ACS signature fallback - holiday notice"'));
  assert.ok(ps.includes('From    = @(' + live.sharedMailboxes.map((b) => '"' + b.addresses[0] + '"').join(", ") + ")"));
  assert.ok(ps.includes('Expiry  = [datetime]"2027-01-01T13:00:00.000Z"'));
  assert.ok(ps.includes("if ($r.Expiry) { $params.ExpiryDate = $r.Expiry }"));
  // 最后一条：新建时优先级最低，文字在团队信息下面
  assert.ok(ps.lastIndexOf("holiday notice") > ps.lastIndexOf("ACS signature fallback - " + live.sharedMailboxes.at(-1).team));
  const html = buildHolidayHtml(live);
  assert.ok(text(html).includes("Office closed: 25, 26 & 28 Dec 2026 and 1 Jan 2027"));
  assert.ok(exchangeMatches(holidayPattern(live), html), "rule must recognise its own notice (no stacking in a thread)");
  // 手机签名里有团队那一行，但不能挡住放假通知
  for (const box of live.sharedMailboxes) assert.ok(!exchangeMatches(holidayPattern(live), mobileSignatureText(box, live).replace(/\n/g, "<br>")));
  const gone = renderRulesScript(withUntil("2000-01-01T00:00:00Z"));
  assert.ok(!gone.includes('Name    = "ACS signature fallback - holiday notice"') && !gone.includes("Expiry  = [datetime]"));
});
