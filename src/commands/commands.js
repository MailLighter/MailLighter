/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, console, localStorage, window, URLSearchParams */

import { t } from "../shared/i18n";
import { sanitizeSelectionHtml, toHtmlFromText } from "../shared/office-helpers";
import { findReplySeparators } from "../shared/reply-detection";
import { addSavings, getSavings } from "../shared/savings-storage";

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

const ECO_MESSAGE_KEY = "maillighter_eco_message";
const ECO_MESSAGE_TEXT_KEY = "maillighter_eco_message_text";

function isEcoMessageEnabled() {
  return localStorage.getItem(ECO_MESSAGE_KEY) === "1";
}

function getEcoMessageText() {
  return localStorage.getItem(ECO_MESSAGE_TEXT_KEY) || t("settings.ecoMessageDefault");
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

async function appendEcoMessage() {
  const htmlBody = await getBodyAsync(Office.CoercionType.Html);
  const safeText = escapeHtml(getEcoMessageText());
  const linkedText = safeText.replace(
    /MailLighter/g,
    '<a href="https://www.maillighter.com" style="color:#1b5e20;">MailLighter</a>'
  );
  const ecoHtml =
    `<div style="margin-top:12px;padding-top:8px;border-top:1px solid #c8e6c9;color:#2e7d32;font-size:13px;">` +
    `${linkedText}</div>`;
  await setBodyAsync(htmlBody + ecoHtml, Office.CoercionType.Html);
}

function notify(message, icon = "Icon.80x80") {
  const item = Office.context.mailbox.item;

  if (!item || !item.notificationMessages) {
    return;
  }

  item.notificationMessages.replaceAsync("MailLighterNotification", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message,
    icon,
    persistent: false,
  });
}

function officeAsync(target, method, unavailableKey, failedKey, ...args) {
  return new Promise((resolve, reject) => {
    if (!target || typeof target[method] !== "function") {
      reject(new Error(t(unavailableKey)));
      return;
    }

    target[method](...args, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
        return;
      }

      const rawMessage = result.error && result.error.message ? result.error.message : "";
      if (rawMessage) {
        console.error(`[MailLighter] ${method} failed:`, rawMessage);
      }
      reject(new Error(t(failedKey)));
    });
  });
}

function getBodyAsync(coercionType) {
  const body = Office.context.mailbox.item && Office.context.mailbox.item.body;
  return officeAsync(
    body,
    "getAsync",
    "commands.errors.bodyReadUnavailable",
    "commands.errors.bodyReadFailed",
    coercionType
  ).then((v) => v || "");
}

function setBodyAsync(content, coercionType) {
  const body = Office.context.mailbox.item && Office.context.mailbox.item.body;
  return officeAsync(
    body,
    "setAsync",
    "commands.errors.bodyWriteUnavailable",
    "commands.errors.bodyWriteFailed",
    content,
    { coercionType }
  );
}

function getAttachmentsAsync() {
  const item = Office.context.mailbox.item;
  return officeAsync(
    item,
    "getAttachmentsAsync",
    "commands.errors.attachmentsUnavailableContext",
    "commands.errors.attachmentsReadFailed"
  ).then((v) => v || []);
}

function removeAttachmentAsync(attachmentId) {
  const item = Office.context.mailbox.item;
  return officeAsync(
    item,
    "removeAttachmentAsync",
    "commands.errors.attachmentsUnavailable",
    "commands.errors.attachmentRemoveFailed",
    attachmentId
  );
}

function calculateImageSize(imgMatches) {
  let totalSize = 0;

  imgMatches.forEach((imgTag) => {
    const dataSizeMatch = imgTag.match(/data-size="?(\d+)"?/i);

    if (dataSizeMatch) {
      totalSize += parseInt(dataSizeMatch[1], 10);
      return;
    }

    const widthMatch = imgTag.match(/width="?(\d+)"?/i);
    const heightMatch = imgTag.match(/height="?(\d+)"?/i);

    if (widthMatch && heightMatch) {
      const width = parseInt(widthMatch[1], 10);
      const height = parseInt(heightMatch[1], 10);
      const estimatedSize = Math.round((width * height) / 2000) * 5120;
      totalSize += estimatedSize;
      return;
    }

    totalSize += 51200;
  });

  return totalSize;
}

function formatFileSize(bytes) {
  if (!bytes || bytes <= 0) {
    return "";
  }

  const kilobytes = bytes / 1024;

  if (kilobytes < 1) {
    return t("units.lessThanOne");
  }

  if (kilobytes < 1024) {
    return `${Math.round(kilobytes * 100) / 100} ${t("units.kilobytes")}`;
  }

  if (kilobytes < 1024 * 1024) {
    return `${Math.round((kilobytes / 1024) * 100) / 100} ${t("units.megabytes")}`;
  }

  return `${Math.round((kilobytes / (1024 * 1024)) * 100) / 100} ${t("units.gigabytes")}`;
}

async function executeWithNotification(
  event,
  worker,
  errorMessage = t("commands.notifications.genericError")
) {
  try {
    const successMessage = await worker();
    notify(successMessage);
  } catch (error) {
    if (error instanceof Error && error.message) {
      console.error("[MailLighter]", error.message);
    }
    notify(errorMessage);
  } finally {
    event.completed();
  }
}

function getSelectedDataAsync(coercionType) {
  const item = Office.context.mailbox.item;
  return officeAsync(
    item,
    "getSelectedDataAsync",
    "commands.errors.selectionUnavailable",
    "commands.errors.selectionReadFailed",
    coercionType
  ).then((v) => (v && v.data ? String(v.data) : ""));
}

async function getSelectedHtmlAsync() {
  let htmlSelection = "";

  try {
    htmlSelection = (await getSelectedDataAsync(Office.CoercionType.Html)).trim();
  } catch {
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
    textSelection = (await getSelectedDataAsync(Office.CoercionType.Text)).trim();
  } catch {
    textSelection = "";
  }

  if (textSelection) {
    return toHtmlFromText(textSelection);
  }

  throw new Error(t("commands.errors.selectionEmpty"));
}

function prependAsync(content, coercionType) {
  const body = Office.context.mailbox.item && Office.context.mailbox.item.body;
  return officeAsync(
    body,
    "prependAsync",
    "commands.errors.bodyWriteUnavailable",
    "commands.errors.bodyWriteFailed",
    content,
    { coercionType }
  );
}

async function keepSelectionOnlyCore() {
  // 1. Capture selection BEFORE reading the full body
  const selectedHtml = await getSelectedHtmlAsync();

  // 2. Read the full body to find the structure
  const fullBody = await getBodyAsync(Office.CoercionType.Html);

  // 3. Find where the quoted/forwarded content starts.
  //    We preserve the user's composing area (typing space, appendonsend, signature)
  //    and replace everything after _MailOriginal with our separator + selection.
  const mailOriginalMatch = fullBody.match(/<a[^>]*name\s*=\s*["']_MailOriginal["'][^>]*>/i);
  const divRplyMatch = fullBody.match(/<div[^>]*\bid\s*=\s*["']divRplyFwdMsg["'][^>]*>/i);

  // Pick the earliest marker
  let cutPoint = -1;
  if (mailOriginalMatch) cutPoint = mailOriginalMatch.index;
  if (divRplyMatch && (cutPoint === -1 || divRplyMatch.index < cutPoint)) {
    cutPoint = divRplyMatch.index;
  }

  const separator = '<hr style="border:none;border-top:1px solid #b5b5b5;margin:8px 0;">';

  if (cutPoint > 0) {
    const userArea = fullBody.substring(0, cutPoint);
    const newBody = userArea + separator + selectedHtml;
    await setBodyAsync(newBody, Office.CoercionType.Html);
  } else {
    const parts = ["<div><br></div>", separator, selectedHtml];
    await setBodyAsync(parts.join(""), Office.CoercionType.Html);
  }

  // 4. Prepend an empty line to place the cursor at the very top.
  //    Works in Old Outlook. In New Outlook the cursor stays at the bottom
  //    (known Office.js API limitation).
  try {
    await prependAsync("<div><br></div>", Office.CoercionType.Html);
  } catch {
    // prependAsync not available in this client
  }

  return t("commands.notifications.keepSelectionDone");
}

async function removeImagesWork() {
  const htmlBody = await getBodyAsync(Office.CoercionType.Html);
  const imgMatches = htmlBody.match(/<img[^>]*>/gi) || [];

  if (imgMatches.length === 0) {
    return { count: 0, sizeText: "", totalBytes: 0 };
  }

  const cleanedHtml = htmlBody.replace(/<img[^>]*>/gi, "");
  await setBodyAsync(cleanedHtml, Office.CoercionType.Html);

  const totalBytes = calculateImageSize(imgMatches);
  return { count: imgMatches.length, sizeText: formatFileSize(totalBytes), totalBytes };
}

async function removeImagesCore() {
  const { count, sizeText, totalBytes } = await removeImagesWork();

  if (count === 0) {
    return t("commands.notifications.imagesNone");
  }

  addSavings("images", totalBytes);

  return sizeText
    ? t("commands.notifications.imagesRemovedWithSize", { count, size: sizeText })
    : t("commands.notifications.imagesRemoved", { count });
}

async function removeAttachmentsWork() {
  const allAttachments = await getAttachmentsAsync();
  // Keep only real file attachments, not inline images (signatures, logos, etc.)
  const attachments = allAttachments.filter((a) => !a.isInline);

  if (attachments.length === 0) {
    return { count: 0, sizeText: "" };
  }

  const totalSize = attachments.reduce((sum, attachment) => {
    const size = typeof attachment.size === "number" ? attachment.size : 0;
    return sum + size;
  }, 0);

  await Promise.all(attachments.map((attachment) => removeAttachmentAsync(attachment.id)));

  return { count: attachments.length, sizeText: formatFileSize(totalSize), totalBytes: totalSize };
}

async function removeAttachmentsCore() {
  const { count, sizeText, totalBytes } = await removeAttachmentsWork();

  if (count === 0) {
    return t("commands.notifications.attachmentsNone");
  }

  addSavings("attachments", totalBytes);

  return sizeText
    ? t("commands.notifications.attachmentsRemovedWithSize", { count, size: sizeText })
    : t("commands.notifications.attachmentsRemoved", { count });
}

async function keepTwoRepliesWork() {
  const htmlBody = await getBodyAsync(Office.CoercionType.Html);
  const separators = findReplySeparators(htmlBody);

  if (separators.length === 0) {
    return { found: 0, cleaned: false };
  }

  if (separators.length <= 2) {
    return { found: separators.length, cleaned: false };
  }

  let cutPoint = separators[2];
  // Never move the cut earlier than the 2nd reply boundary — that content
  // is part of the reply we want to keep.
  const floor = separators[1];
  const before = htmlBody.substring(0, cutPoint);

  const lastHr = before.lastIndexOf("<hr");
  if (lastHr >= 0 && cutPoint - lastHr < 500 && lastHr >= floor) {
    cutPoint = lastHr;
  } else {
    // Detect "---separator text---" reply/forward headers (Outlook, all locales).
    // e.g. "-----Message d'origine-----", "-----Original Message-----",
    //      "---------- Message original ----------", "--- Forwarded message ---"
    const dashSepRe =
      /[-\u2010-\u2014]{3,}[ \t\xa0]*[^-\u2010-\u2014\n\r<]{3,60}[ \t\xa0]*[-\u2010-\u2014]{3,}/g;
    let lastDashSepIdx = -1;
    let dashMatch;
    while ((dashMatch = dashSepRe.exec(before)) !== null) {
      if (dashMatch.index >= floor) {
        lastDashSepIdx = dashMatch.index;
      }
    }
    if (lastDashSepIdx >= 0 && cutPoint - lastDashSepIdx < 1000) {
      const tagStart = before.lastIndexOf("<", lastDashSepIdx);
      cutPoint = tagStart >= 0 && tagStart >= floor ? tagStart : lastDashSepIdx;
    }
  }

  const savedBytes = htmlBody.length - cutPoint;
  await setBodyAsync(htmlBody.substring(0, cutPoint), Office.CoercionType.Html);
  return { found: separators.length, cleaned: true, savedBytes };
}

async function keepTwoRepliesCore() {
  const { found, cleaned, savedBytes } = await keepTwoRepliesWork();

  if (found === 0) {
    return t("commands.notifications.repliesNone");
  }

  if (!cleaned) {
    return t("commands.notifications.repliesNoChange", { count: found });
  }

  if (isEcoMessageEnabled()) {
    await appendEcoMessage();
  }

  addSavings("replies", savedBytes);

  const sizeText = savedBytes ? formatFileSize(savedBytes) : "";
  return sizeText
    ? t("commands.notifications.repliesCleanedWithSize", { count: found, size: sizeText })
    : t("commands.notifications.repliesCleaned", { count: found });
}

function formatCleanAllPart(prefix, count, sizeText) {
  const colon = t("units.colon");
  if (!count) return prefix + colon + "0";
  return sizeText ? `${prefix}${colon}${count} (${sizeText})` : `${prefix}${colon}${count}`;
}

async function cleanAllCore() {
  const parts = [];
  let totalBytes = 0;

  try {
    const img = await removeImagesWork();
    parts.push(
      formatCleanAllPart(t("commands.notifications.cleanAllImagesPrefix"), img.count, img.sizeText)
    );
    totalBytes += img.totalBytes || 0;
    addSavings("images", img.totalBytes || 0);
  } catch (error) {
    parts.push(`${t("commands.notifications.cleanAllImagesPrefix")}: ${error.message}`);
  }

  try {
    const att = await removeAttachmentsWork();
    parts.push(
      formatCleanAllPart(
        t("commands.notifications.cleanAllAttachmentsPrefix"),
        att.count,
        att.sizeText
      )
    );
    totalBytes += att.totalBytes || 0;
    addSavings("attachments", att.totalBytes || 0);
  } catch (error) {
    parts.push(`${t("commands.notifications.cleanAllAttachmentsPrefix")}: ${error.message}`);
  }

  try {
    const rep = await keepTwoRepliesWork();
    const prefix = t("commands.notifications.cleanAllRepliesPrefix");
    const colon = t("units.colon");
    if (rep.found === 0) {
      parts.push(prefix + colon + "0");
    } else if (rep.cleaned) {
      const repSizeText = rep.savedBytes ? formatFileSize(rep.savedBytes) : "";
      parts.push(
        repSizeText
          ? `${prefix}${colon}${rep.found} → 2 (${repSizeText})`
          : `${prefix}${colon}${rep.found} → 2`
      );
      totalBytes += rep.savedBytes || 0;
      addSavings("replies", rep.savedBytes || 0);
      if (isEcoMessageEnabled()) {
        await appendEcoMessage();
      }
    } else {
      parts.push(`${prefix}${colon}${rep.found}`);
    }
  } catch (error) {
    parts.push(`${t("commands.notifications.cleanAllRepliesPrefix")}: ${error.message}`);
  }

  const totalText =
    totalBytes > 0
      ? t("commands.notifications.cleanAllTotal", { size: formatFileSize(totalBytes) })
      : "";

  return t("commands.notifications.cleanAllDone", { details: parts.join(" | "), total: totalText });
}

function openSettingsCore(event) {
  const ecoEnabled = isEcoMessageEnabled();
  const savings = getSavings();
  const params = new URLSearchParams({
    ecoMessage: ecoEnabled ? "1" : "0",
    ecoText: getEcoMessageText(),
    savImages: savings.images,
    savReplies: savings.replies,
    savAttachments: savings.attachments,
    savTotal: savings.total,
  });
  const settingsUrl = `${window.location.origin}/settings.html?${params}`;

  Office.context.ui.displayDialogAsync(settingsUrl, { height: 65, width: 40 }, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      notify(t("commands.notifications.cannotOpenSettings"));
      event.completed();
      return;
    }

    const dialog = result.value;
    let completed = false;
    const finish = () => {
      if (completed) return;
      completed = true;
      event.completed();
    };

    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
      let data;
      try {
        data = JSON.parse(arg.message);
      } catch {
        return;
      }
      if (typeof data.ecoMessageEnabled !== "undefined") {
        localStorage.setItem(ECO_MESSAGE_KEY, data.ecoMessageEnabled ? "1" : "0");
      }
      if (typeof data.ecoMessageText !== "undefined") {
        localStorage.setItem(ECO_MESSAGE_TEXT_KEY, data.ecoMessageText);
      }
      if (data.action === "close") {
        try {
          dialog.close();
        } catch {
          // dialog handle stale — dialog closed itself via window.close()
        }
        // Programmatic close via dialog.close() does not fire
        // DialogEventReceived, so finish the command here.
        finish();
      }
    });

    // Fires on user-initiated close (clicking X) or runtime errors.  We
    // keep the command alive until then to prevent new Outlook / Web from
    // tearing down the commands runtime and the dialog with it.
    dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
      finish();
    });
  });
}

// Register all commands with Office.
[
  ["removeImagesCommand", removeImagesCore, "cannotRemoveImages"],
  ["removeAttachmentsCommand", removeAttachmentsCore, "cannotRemoveAttachments"],
  ["keepTwoRepliesCommand", keepTwoRepliesCore, "cannotKeepReplies"],
  ["cleanAllCommand", cleanAllCore, "cannotCleanAll"],
  ["keepSelectionOnlyCommand", keepSelectionOnlyCore, "cannotKeepSelection"],
].forEach(([name, core, errorKey]) => {
  Office.actions.associate(name, (event) => {
    executeWithNotification(event, core, t(`commands.notifications.${errorKey}`));
  });
});

Office.actions.associate("openSettingsCommand", openSettingsCore);
