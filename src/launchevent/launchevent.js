/* global Office */
/**
 * 事件入口：新建 / 回复 / 转发邮件时自动插入签名。
 * 经典 Outlook (Windows) 直接加载本文件（JS-only 运行时）；
 * 新版 Outlook、网页版、Mac、iOS、Android 通过 commands.html 加载。
 */
import { loadProfile, decideVariant, applySignature, log } from "../shared/office.js";

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

async function onNewMessageComposeHandler(event) {
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
    const decision = await decideVariant(res.profile, res.token);
    log("decision", decision.variant, decision.reason, decision.composeType);
    await applySignature(decision.variant, res.profile);
    if (res.profile.source === "fallback") notifySignIn();
  } catch (e) {
    log("handler error", e && (e.message || e));
  } finally {
    clearTimeout(guard);
    finish();
  }
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
