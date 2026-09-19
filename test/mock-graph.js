// 测试用：替换 graph.js，模拟 Graph 返回
export const SCOPES = [];
export function isNaaSupported() { return true; }
export async function getToken() { if (globalThis.__MOCK.tokenFail) throw new Error("no token"); return "tok"; }
export async function fetchMe() { return globalThis.__MOCK.me; }
export function conversationIdVariants(id) { return [id]; }
export async function hasSentInConversation(t, id) { globalThis.__MOCK.checkedConv = id; return !!globalThis.__MOCK.sent; }
