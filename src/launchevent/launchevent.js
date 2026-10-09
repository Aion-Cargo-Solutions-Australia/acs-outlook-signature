/* global Office */
/**
 * 事件入口：新建 / 回复 / 转发邮件时自动插入签名；切换发件人 (From) 时换成对应的签名。
 * 经典 Outlook (Windows) 直接加载本文件（JS-only 运行时）；
 * 新版 Outlook、网页版、Mac、iOS、Android 通过 commands.html 加载。
 */
import { loadProfile, resolveSender, decideVariant, applySignature, log } from "../shared/office.js";

Office.onReady(() => {});

const SIGN_IN_NOTICE_ID = "acs_sig_signin";

function notifySignIn() {
  try {
    Office.context.mailbox.item.notificationMessages.addAsync(SIGN_IN_NOTICE_ID, {
      type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
      message: "ACS 签名无法读取你的联系方式，请打开 ACS Signature 面板登录一次。",
      icon: "Icon.16x16",
      actions: [
        {
          actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
          actionText: "打开签名面板",
          commandId: "ACS_SigPaneButton",
          contextData: "{}",
        },
      ],
    });
  } catch (e) {
    // 手机端不支持带按钮的通知，忽略
  }
}

/** @param {boolean} reapply 这封邮件里已经插入过签名（切换发件人），这次是替换 */
async function insertSignature(event, reapply) {
  let done = false;
  const finish = () => {
    if (!done) {
      done = true;
      event.completed();
    }
  };
  // 保险：最迟 45 秒结束（手机端上限 60 秒）
  const guard = setTimeout(finish, 45000);
  try {
    const res = await loadProfile(false);
    const sender = await resolveSender(res.profile);
    const decision = await decideVariant(sender, res.token, { reapply: reapply });
    log("decision", decision.variant, decision.reason, decision.composeType, sender.shared ? "team " + sender.email : "personal");
    await applySignature(decision.variant, sender, { reapply: reapply });
    if (res.profile.source === "fallback") notifySignIn();
  } catch (e) {
    log("handler error", e && (e.message || e));
  } finally {
    clearTimeout(guard);
    finish();
  }
}

function onNewMessageComposeHandler(event) {
  return insertSignature(event, false);
}

// 在个人邮箱和共享邮箱之间切换发件人时触发（重新选了同一个发件人也会触发）
function onMessageFromChangedHandler(event) {
  return insertSignature(event, true);
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onMessageFromChangedHandler", onMessageFromChangedHandler);
