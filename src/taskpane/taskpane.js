/* global Office, document */
import { loadProfile, decideVariant, applySignature, buildHtml } from "../shared/office.js";

let state = { profile: null, token: null, variant: "full" };

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s == null ? "" : s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

function renderProfile(p) {
  const rows = [
    ["姓名", p.displayName], ["职位", p.jobTitle], ["邮箱", p.email], ["手机", p.mobile],
    ["电话", p.phone], ["分机", p.ext], ["地址", p.address || "（使用公司地址）"],
  ];
  $("profileTable").innerHTML = rows
    .map((r) => "<tr><td>" + r[0] + "</td><td>" + (r[1] ? esc(r[1]) : '<span class="empty">未填写</span>') + "</td></tr>")
    .join("");
}

function renderPreview() {
  // 预览统一使用外链图片（内嵌 cid 图片只能在邮件里显示）
  $("preview").innerHTML = buildHtml(state.variant, state.profile, "link");
  document.querySelectorAll(".tab").forEach((b) => b.classList.toggle("active", b.dataset.variant === state.variant));
}

async function load(allowPopup) {
  $("statusText").textContent = "正在读取你的资料…";
  const res = await loadProfile(allowPopup);
  state.profile = res.profile;
  state.token = res.token;
  if (res.token) {
    $("statusText").innerHTML = '<span class="ok">✔ 已连接 Entra ID，签名会自动插入。</span>';
    $("btnSignIn").hidden = true;
  } else {
    $("statusText").innerHTML = '<span class="warn">未能读取 Entra 资料' + (res.error ? "（" + esc(res.error.message || res.error) + "）" : "") + "。请点击下方按钮登录授权一次。</span>";
    $("btnSignIn").hidden = false;
  }
  renderProfile(state.profile);
  try {
    const d = await decideVariant(state.profile, state.token);
    state.variant = d.variant;
    $("decision").textContent = "自动判断：" + (d.variant === "full" ? "完整签名" : "精简签名") + "（" + d.reason + "）";
  } catch (e) {
    $("decision").textContent = "";
  }
  renderPreview();
}

Office.onReady(() => {
  document.querySelectorAll(".tab").forEach((b) =>
    b.addEventListener("click", () => {
      state.variant = b.dataset.variant;
      renderPreview();
    })
  );
  $("btnSignIn").addEventListener("click", () => load(true));
  $("btnRefresh").addEventListener("click", () => load(true));
  $("btnInsert").addEventListener("click", async () => {
    $("msg").textContent = "正在插入…";
    try {
      await applySignature(state.variant, state.profile);
      $("msg").textContent = "已插入。";
    } catch (e) {
      $("msg").textContent = "插入失败：" + (e.message || e);
    }
  });
  load(false);
});
