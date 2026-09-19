/* global Office */
// 用 roamingSettings 缓存用户资料（跨设备同步，且事件运行时可用；localStorage 在经典 Outlook 事件运行时里不可用）
const KEY = "acsSigProfile";
const MAX_AGE_MS = 12 * 60 * 60 * 1000;

export function readCachedProfile() {
  try {
    const raw = Office.context.roamingSettings.get(KEY);
    if (!raw) return null;
    const obj = typeof raw === "string" ? JSON.parse(raw) : raw;
    return obj && obj.profile ? obj : null;
  } catch (e) {
    return null;
  }
}

export function isFresh(entry) {
  return !!entry && Date.now() - entry.savedAt < MAX_AGE_MS;
}

export function saveCachedProfile(profile) {
  return new Promise((resolve) => {
    try {
      Office.context.roamingSettings.set(KEY, JSON.stringify({ savedAt: Date.now(), profile: profile }));
      Office.context.roamingSettings.saveAsync(() => resolve());
    } catch (e) {
      resolve();
    }
  });
}
