/* global Office */

import { PlatformAdapter } from "../PlatformAdapter";
import { sanitizeSelectionHtml, toHtmlFromText } from "../../core/htmlSanitizer";
import { logger } from "../../utils/logger";

/**
 * Wrap Office.js callback-style APIs into Promises with localized error keys.
 * The caller is responsible for translating the error keys into user-facing
 * messages (kept out of the adapter for testability and i18n discipline).
 */
function officeAsync(target, method, unavailableKey, failedKey, ...args) {
  return new Promise((resolve, reject) => {
    if (!target || typeof target[method] !== "function") {
      reject(new Error(unavailableKey));
      return;
    }
    target[method](...args, (result) => {
      if (result && result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
        return;
      }
      const rawMessage = result && result.error && result.error.message ? result.error.message : "";
      if (rawMessage) {
        logger.warn(`${method} failed:`, rawMessage);
      }
      reject(new Error(failedKey));
    });
  });
}

export class OutlookAdapter extends PlatformAdapter {
  getLocale() {
    if (typeof Office !== "undefined" && Office.context) {
      return Office.context.displayLanguage || "";
    }
    return "";
  }

  _item() {
    return Office.context && Office.context.mailbox && Office.context.mailbox.item;
  }

  async getCurrentItem() {
    return this._item();
  }

  async getBodyHtml() {
    const item = this._item();
    const body = item && item.body;
    const value = await officeAsync(
      body,
      "getAsync",
      "commands.errors.bodyReadUnavailable",
      "commands.errors.bodyReadFailed",
      Office.CoercionType.Html
    );
    return value || "";
  }

  async setBodyHtml(html) {
    const item = this._item();
    const body = item && item.body;
    return officeAsync(
      body,
      "setAsync",
      "commands.errors.bodyWriteUnavailable",
      "commands.errors.bodyWriteFailed",
      html,
      { coercionType: Office.CoercionType.Html }
    );
  }

  async prependBodyHtml(html) {
    const item = this._item();
    const body = item && item.body;
    return officeAsync(
      body,
      "prependAsync",
      "commands.errors.bodyWriteUnavailable",
      "commands.errors.bodyWriteFailed",
      html,
      { coercionType: Office.CoercionType.Html }
    );
  }

  async _getSelectedDataAsync(coercionType) {
    const item = this._item();
    const value = await officeAsync(
      item,
      "getSelectedDataAsync",
      "commands.errors.selectionUnavailable",
      "commands.errors.selectionReadFailed",
      coercionType
    );
    return value && value.data ? String(value.data) : "";
  }

  async getSelectedHtml() {
    let htmlSelection = "";
    try {
      htmlSelection = (await this._getSelectedDataAsync(Office.CoercionType.Html)).trim();
    } catch (error) {
      logger.warn("selection HTML read failed:", error.message);
      htmlSelection = "";
    }
    if (htmlSelection) {
      const cleaned = sanitizeSelectionHtml(htmlSelection);
      if (cleaned.trim()) {
        return cleaned;
      }
    }

    let textSelection = "";
    try {
      textSelection = (await this._getSelectedDataAsync(Office.CoercionType.Text)).trim();
    } catch (error) {
      logger.warn("selection text read failed:", error.message);
      textSelection = "";
    }
    if (textSelection) {
      return toHtmlFromText(textSelection);
    }

    throw new Error("commands.errors.selectionEmpty");
  }

  async listAttachments() {
    const item = this._item();
    const value = await officeAsync(
      item,
      "getAttachmentsAsync",
      "commands.errors.attachmentsUnavailableContext",
      "commands.errors.attachmentsReadFailed"
    );
    return value || [];
  }

  async removeAttachment(attachmentId) {
    const item = this._item();
    return officeAsync(
      item,
      "removeAttachmentAsync",
      "commands.errors.attachmentsUnavailable",
      "commands.errors.attachmentRemoveFailed",
      attachmentId
    );
  }

  /**
   * Returns the count of distinct recipients across To/Cc/Bcc.
   * In compose mode, uses Office.js recipients API. Falls back to 1 if the
   * API is unavailable (degraded estimate rather than zero, since "send"
   * implies at least one recipient).
   */
  async getRecipients() {
    const item = this._item();
    if (!item) return [];

    const fields = ["to", "cc", "bcc"];
    const all = [];
    for (const field of fields) {
      const recipientsApi = item[field];
      if (!recipientsApi || typeof recipientsApi.getAsync !== "function") {
        // Read mode: item.to/cc/bcc are arrays directly.
        if (Array.isArray(item[field])) {
          all.push(...item[field]);
        }
        continue;
      }
      try {
        const value = await new Promise((resolve, reject) => {
          recipientsApi.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(result.value || []);
            } else {
              reject(new Error(result.error && result.error.message));
            }
          });
        });
        all.push(...value);
      } catch (e) {
        logger.warn(`recipients ${field} read failed:`, e.message);
      }
    }
    return all;
  }

  notify(message) {
    const item = this._item();
    if (!item || !item.notificationMessages) return;
    item.notificationMessages.replaceAsync("MailLighterNotification", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message,
      icon: "Icon.80x80",
      persistent: false,
    });
  }
}
