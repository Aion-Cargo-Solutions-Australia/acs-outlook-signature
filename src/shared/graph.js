/* global Office, fetch */
import { createNestablePublicClientApplication } from "@azure/msal-browser";
import { ADDIN } from "../generated/addin.js";
import { GRAPH_SELECT } from "./profile.js";

// 最小权限：User.Read 读取本人资料；Mail.ReadBasic 只读邮件元数据（不含正文），用于判断会话里是否已回复过
export const SCOPES = ["User.Read", "Mail.ReadBasic"];

let pcaPromise = null;

function getPca() {
  if (!pcaPromise) {
    pcaPromise = createNestablePublicClientApplication({
      auth: {
        clientId: ADDIN.clientId,
        authority: "https://login.microsoftonline.com/" + (ADDIN.tenantId || "organizations"),
      },
    });
  }
  return pcaPromise;
}

export function isNaaSupported() {
  try {
    return !!(Office.context.requirements && Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"));
  } catch (e) {
    return false;
  }
}

/** 静默取 token（事件处理器里只能静默）。allowPopup=true 仅在任务窗格中使用。 */
export async function getToken(allowPopup) {
  const pca = await getPca();
  const request = { scopes: SCOPES };
  try {
    const r = await pca.acquireTokenSilent(request);
    return r.accessToken;
  } catch (silentError) {
    if (!allowPopup) throw silentError;
    const r = await pca.acquireTokenPopup(request);
    return r.accessToken;
  }
}

async function graphGet(token, path) {
  const res = await fetch("https://graph.microsoft.com/v1.0" + path, {
    headers: { Authorization: "Bearer " + token, ConsistencyLevel: "eventual" },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error("Graph " + res.status + ": " + text.slice(0, 300));
  }
  return res.json();
}

export function fetchMe(token) {
  return graphGet(token, "/me?$select=" + GRAPH_SELECT);
}

/** Office.js 的 conversationId 可能是 EWS 格式（含 + /），Graph 用 REST 格式（- _），两种都查。 */
export function conversationIdVariants(id) {
  if (!id) return [];
  const rest = id.replace(/\+/g, "-").replace(/\//g, "_");
  const ews = id.replace(/-/g, "+").replace(/_/g, "/");
  return Array.from(new Set([id, rest, ews]));
}

/** 本人是否已经在该会话中发送过邮件（查“已发送邮件”文件夹） */
export async function hasSentInConversation(token, conversationId) {
  const ids = conversationIdVariants(conversationId);
  if (!ids.length) return false;
  const filter = ids.map((v) => "conversationId eq '" + v.replace(/'/g, "''") + "'").join(" or ");
  const data = await graphGet(
    token,
    "/me/mailFolders/SentItems/messages?$select=id&$top=1&$filter=" + encodeURIComponent(filter)
  );
  return !!(data.value && data.value.length);
}
