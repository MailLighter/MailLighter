/* global Office, document, setTimeout */

import { t } from "../shared/i18n";
import {
  sanitizeHtml,
  toHtmlFromText,
  displayFormWithFallback,
  getForwardSubject,
} from "../shared/office-helpers";

Office.onReady(() => {
  document.getElementById("btnReply").onclick = () => doPartialAction("reply");
  document.getElementById("btnReplyAll").onclick = () => doPartialAction("replyAll");
  document.getElementById("btnForward").onclick = () => doPartialAction("forward");
  localizeUI();
});

function localizeUI() {
  document.getElementById("hint").textContent = t("taskpane.hint");
  document.getElementById("lblReply").textContent = t("taskpane.reply");
  document.getElementById("lblReplyAll").textContent = t("taskpane.replyAll");
  document.getElementById("lblForward").textContent = t("taskpane.forward");
}

function showStatus(message, type) {
  const el = document.getElementById("status");
  el.textContent = message;
  el.className = "status " + type;
  if (type !== "error") {
    setTimeout(() => {
      el.textContent = "";
      el.className = "status";
    }, 4000);
  }
}


function getSelectedHtml() {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;

    if (!item || typeof item.getSelectedDataAsync !== "function") {
      resolve("");
      return;
    }

    item.getSelectedDataAsync(Office.CoercionType.Html, (result) => {
      if (
        result.status === Office.AsyncResultStatus.Succeeded &&
        result.value &&
        result.value.data &&
        result.value.data.trim()
      ) {
        const cleaned = sanitizeHtml(result.value.data.trim());

        if (cleaned.trim()) {
          resolve(cleaned);
          return;
        }
      }

      item.getSelectedDataAsync(Office.CoercionType.Text, (textResult) => {
        if (
          textResult.status === Office.AsyncResultStatus.Succeeded &&
          textResult.value &&
          textResult.value.data &&
          textResult.value.data.trim()
        ) {
          resolve(toHtmlFromText(textResult.value.data.trim()));
          return;
        }

        resolve("");
      });
    });
  });
}


async function doPartialAction(type) {
  const item = Office.context.mailbox.item;

  if (!item) {
    showStatus(t("commands.errors.selectionUnavailable"), "error");
    return;
  }

  try {
    const html = await getSelectedHtml();
    const noSelection = !html;

    if (type === "reply") {
      if (typeof item.displayReplyForm !== "function") {
        showStatus(t("commands.errors.replyFormUnavailable"), "error");
        return;
      }
      displayFormWithFallback(item.displayReplyForm.bind(item), html);
    } else if (type === "replyAll") {
      if (typeof item.displayReplyAllForm !== "function") {
        showStatus(t("commands.errors.replyAllFormUnavailable"), "error");
        return;
      }
      displayFormWithFallback(item.displayReplyAllForm.bind(item), html);
    } else if (type === "forward") {
      if (typeof item.displayForwardForm === "function") {
        displayFormWithFallback(item.displayForwardForm.bind(item), html);
      } else if (
        Office.context.mailbox &&
        typeof Office.context.mailbox.displayNewMessageForm === "function"
      ) {
        const payload = { subject: getForwardSubject(item.subject) };
        if (html) payload.htmlBody = html;
        Office.context.mailbox.displayNewMessageForm(payload);
      } else {
        showStatus(t("commands.errors.forwardFormUnavailable"), "error");
        return;
      }
    }

    const msgKey =
      type === "reply"
        ? "partialReplyOpened"
        : type === "replyAll"
          ? "partialReplyAllOpened"
          : "partialForwardOpened";

    const msg = t(`commands.notifications.${msgKey}`);
    showStatus(
      noSelection ? `${msg} ${t("commands.notifications.partialOpenedWithoutSelection")}` : msg,
      noSelection ? "info" : "success"
    );
  } catch (error) {
    showStatus(error.message || t("commands.notifications.genericError"), "error");
  }
}
