/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

import { t } from "../shared/i18n";

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

function notify(message, icon = "Icon.80x80") {
  const item = Office.context.mailbox.item;

  if (!item || !item.notificationMessages) {
    return;
  }

  if (typeof item.notificationMessages.removeAsync === "function") {
    // Remove legacy sample notification if it exists from older builds.
    item.notificationMessages.removeAsync("ActionPerformanceNotification");
  }

  item.notificationMessages.replaceAsync("MailLighterNotification", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message,
    icon,
    persistent: false,
  });
}

function getBodyAsync(coercionType) {
  const item = Office.context.mailbox.item;

  return new Promise((resolve, reject) => {
    if (!item || !item.body || typeof item.body.getAsync !== "function") {
      reject(new Error(t("commands.errors.bodyReadUnavailable")));
      return;
    }

    item.body.getAsync(coercionType, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
        return;
      }

      reject(
        new Error(
          result.error && result.error.message
            ? result.error.message
            : t("commands.errors.bodyReadFailed")
        )
      );
    });
  });
}

function setBodyAsync(content, coercionType) {
  const item = Office.context.mailbox.item;

  return new Promise((resolve, reject) => {
    if (!item || !item.body || typeof item.body.setAsync !== "function") {
      reject(new Error(t("commands.errors.bodyWriteUnavailable")));
      return;
    }

    item.body.setAsync(content, { coercionType }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
        return;
      }

      reject(
        new Error(
          result.error && result.error.message
            ? result.error.message
            : t("commands.errors.bodyWriteFailed")
        )
      );
    });
  });
}

function getAttachmentsAsync() {
  const item = Office.context.mailbox.item;

  return new Promise((resolve, reject) => {
    if (!item || typeof item.getAttachmentsAsync !== "function") {
      reject(new Error(t("commands.errors.attachmentsUnavailableContext")));
      return;
    }

    item.getAttachmentsAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || []);
        return;
      }

      reject(
        new Error(
          result.error && result.error.message
            ? result.error.message
            : t("commands.errors.attachmentsReadFailed")
        )
      );
    });
  });
}

function removeAttachmentAsync(attachmentId) {
  const item = Office.context.mailbox.item;

  return new Promise((resolve, reject) => {
    if (!item || typeof item.removeAttachmentAsync !== "function") {
      reject(new Error(t("commands.errors.attachmentsUnavailable")));
      return;
    }

    item.removeAttachmentAsync(attachmentId, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
        return;
      }

      reject(
        new Error(
          result.error && result.error.message
            ? result.error.message
            : t("commands.errors.attachmentRemoveFailed")
        )
      );
    });
  });
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
    const details = error instanceof Error && error.message ? ` (${error.message})` : "";
    notify(`${errorMessage}${details}`);
  } finally {
    event.completed();
  }
}

const MAX_REPLY_BODY_LENGTH = 30000;

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getSelectedDataAsync(coercionType) {
  const item = Office.context.mailbox.item;

  return new Promise((resolve, reject) => {
    if (!item || typeof item.getSelectedDataAsync !== "function") {
      reject(new Error(t("commands.errors.selectionUnavailable")));
      return;
    }

    item.getSelectedDataAsync(coercionType, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(
          new Error(
            result.error && result.error.message
              ? result.error.message
              : t("commands.errors.selectionReadFailed")
          )
        );
        return;
      }

      const selection = result.value && result.value.data ? String(result.value.data) : "";

      resolve(selection);
    });
  });
}

async function getSelectedHtmlAsync() {
  let htmlSelection = "";

  try {
    htmlSelection = (await getSelectedDataAsync(Office.CoercionType.Html)).trim();
  } catch {
    htmlSelection = "";
  }

  if (htmlSelection) {
    const sanitizedSelection = htmlSelection
      .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "")
      .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, "")
      .replace(/<!--[\s\S]*?-->/g, "");

    if (sanitizedSelection.trim()) {
      return sanitizedSelection;
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

function toHtmlFromText(text) {
  return `<div style="white-space: pre-wrap;">${escapeHtml(text || "")}</div>`;
}

function fitHtmlForReply(html) {
  const normalized = String(html || "").trim();

  if (normalized.length <= MAX_REPLY_BODY_LENGTH) {
    return { html: normalized, trimmed: false };
  }

  const textCandidate = normalized
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const truncatedText = textCandidate.slice(0, MAX_REPLY_BODY_LENGTH / 2);

  return {
    html: toHtmlFromText(`${truncatedText}...`),
    trimmed: true,
  };
}

async function getPartialHtmlAsync() {
  try {
    const selectedHtml = await getSelectedHtmlAsync();
    const fitted = fitHtmlForReply(selectedHtml);

    return {
      html: fitted.html,
      openedWithoutSelection: false,
      wasTrimmed: fitted.trimmed,
    };
  } catch {
    return {
      html: "",
      openedWithoutSelection: true,
      wasTrimmed: false,
    };
  }
}

function buildPartialSuccessMessage(baseMessage, options) {
  const notes = [];

  if (options.openedWithoutSelection) {
    notes.push(t("commands.notifications.partialOpenedWithoutSelection"));
  }

  if (options.wasTrimmed) {
    notes.push(t("commands.notifications.partialContentTrimmed"));
  }

  return notes.length > 0 ? `${baseMessage} ${notes.join(" ")}` : baseMessage;
}

function displayFormWithFallback(displayFn, htmlBody) {
  const errors = [];

  const tryInvoke = (argProvided, arg) => {
    try {
      if (argProvided) {
        displayFn(arg);
      } else {
        displayFn();
      }

      return true;
    } catch (error) {
      errors.push(error);
      return false;
    }
  };

  if (htmlBody) {
    if (tryInvoke(true, { htmlBody })) {
      return;
    }

    if (tryInvoke(true, htmlBody)) {
      return;
    }
  }

  if (tryInvoke(false)) {
    return;
  }

  const lastError = errors.length > 0 ? errors[errors.length - 1] : null;

  if (lastError instanceof Error) {
    throw lastError;
  }

  throw new Error(t("commands.notifications.genericError"));
}

function openReplyFormAsync(replyAll, htmlBody) {
  const item = Office.context.mailbox.item;

  return new Promise((resolve, reject) => {
    if (!item) {
      reject(new Error(t("commands.errors.replyFormUnavailable")));
      return;
    }

    const displayFn = replyAll ? item.displayReplyAllForm : item.displayReplyForm;
    const unavailableError = replyAll
      ? t("commands.errors.replyAllFormUnavailable")
      : t("commands.errors.replyFormUnavailable");

    if (typeof displayFn !== "function") {
      reject(new Error(unavailableError));
      return;
    }

    try {
      displayFormWithFallback(displayFn.bind(item), htmlBody);
      resolve();
    } catch (error) {
      reject(error instanceof Error ? error : new Error(String(error)));
    }
  });
}

function getForwardSubject(subject) {
  const sourceSubject = String(subject || "").trim();
  const prefix = t("commands.notifications.forwardPrefix");

  if (!sourceSubject) {
    return prefix;
  }

  return new RegExp(`^${prefix}\\s*`, "i").test(sourceSubject)
    ? sourceSubject
    : `${prefix} ${sourceSubject}`;
}

function openForwardFormAsync(htmlBody) {
  const mailbox = Office.context.mailbox;
  const item = mailbox ? mailbox.item : null;

  return new Promise((resolve, reject) => {
    try {
      if (item && typeof item.displayForwardForm === "function") {
        displayFormWithFallback(item.displayForwardForm.bind(item), htmlBody);
        resolve();
        return;
      }

      if (!mailbox || typeof mailbox.displayNewMessageForm !== "function") {
        reject(new Error(t("commands.errors.forwardFormUnavailable")));
        return;
      }

      const formPayload = {
        subject: getForwardSubject(item && item.subject ? item.subject : ""),
      };

      if (htmlBody) {
        formPayload.htmlBody = htmlBody;
      }

      mailbox.displayNewMessageForm(formPayload);
      resolve();
    } catch (error) {
      reject(error instanceof Error ? error : new Error(String(error)));
    }
  });
}

async function partialReplyCore() {
  const partial = await getPartialHtmlAsync();
  await openReplyFormAsync(false, partial.html);
  return buildPartialSuccessMessage(t("commands.notifications.partialReplyOpened"), partial);
}

async function partialReplyAllCore() {
  const partial = await getPartialHtmlAsync();
  await openReplyFormAsync(true, partial.html);
  return buildPartialSuccessMessage(t("commands.notifications.partialReplyAllOpened"), partial);
}

async function partialForwardCore() {
  const partial = await getPartialHtmlAsync();
  await openForwardFormAsync(partial.html);
  return buildPartialSuccessMessage(t("commands.notifications.partialForwardOpened"), partial);
}

async function removeImagesCore() {
  const htmlBody = await getBodyAsync(Office.CoercionType.Html);
  const imgMatches = htmlBody.match(/<img[^>]*>/gi) || [];

  if (imgMatches.length === 0) {
    return t("commands.notifications.imagesNone");
  }

  const cleanedHtml = htmlBody.replace(/<img[^>]*>/gi, "");
  await setBodyAsync(cleanedHtml, Office.CoercionType.Html);

  const totalSize = calculateImageSize(imgMatches);
  const sizeText = formatFileSize(totalSize);

  return sizeText
    ? t("commands.notifications.imagesRemovedWithSize", {
        count: imgMatches.length,
        size: sizeText,
      })
    : t("commands.notifications.imagesRemoved", { count: imgMatches.length });
}

async function removeAttachmentsCore() {
  const attachments = await getAttachmentsAsync();

  if (attachments.length === 0) {
    return t("commands.notifications.attachmentsNone");
  }

  const totalSize = attachments.reduce((sum, attachment) => {
    const size = typeof attachment.size === "number" ? attachment.size : 0;
    return sum + size;
  }, 0);

  await Promise.all(attachments.map((attachment) => removeAttachmentAsync(attachment.id)));

  const sizeText = formatFileSize(totalSize);

  return sizeText
    ? t("commands.notifications.attachmentsRemovedWithSize", {
        count: attachments.length,
        size: sizeText,
      })
    : t("commands.notifications.attachmentsRemoved", { count: attachments.length });
}

async function keepTwoRepliesCore() {
  const textBody = await getBodyAsync(Office.CoercionType.Text);
  const deMatches = [];
  const deRegex = /De\s*:/gi;
  let match;

  while ((match = deRegex.exec(textBody)) !== null) {
    deMatches.push(match.index);
  }

  if (deMatches.length === 0) {
    return t("commands.notifications.repliesNone");
  }

  if (deMatches.length <= 2) {
    return t("commands.notifications.repliesNoChange", { count: deMatches.length });
  }

  const segments = [];

  for (let index = 0; index < deMatches.length; index += 1) {
    const start = deMatches[index];
    const end = index + 1 < deMatches.length ? deMatches[index + 1] : textBody.length;
    segments.push(textBody.substring(start, end));
  }

  let cleanedText = textBody.substring(0, deMatches[0]);
  cleanedText += segments[0] || "";
  cleanedText += segments[1] || "";

  await setBodyAsync(cleanedText, Office.CoercionType.Text);
  return t("commands.notifications.repliesCleaned", { count: deMatches.length });
}

async function cleanAllCore() {
  const messages = [];

  try {
    messages.push(await removeImagesCore());
  } catch (error) {
    messages.push(`${t("commands.notifications.cleanAllImagesPrefix")}: ${error.message}`);
  }

  try {
    messages.push(await removeAttachmentsCore());
  } catch (error) {
    messages.push(`${t("commands.notifications.cleanAllAttachmentsPrefix")}: ${error.message}`);
  }

  try {
    messages.push(await keepTwoRepliesCore());
  } catch (error) {
    messages.push(`${t("commands.notifications.cleanAllRepliesPrefix")}: ${error.message}`);
  }

  return t("commands.notifications.cleanAllDone", { details: messages.join(" | ") });
}

function removeImagesCommand(event) {
  executeWithNotification(event, removeImagesCore, t("commands.notifications.cannotRemoveImages"));
}

function removeAttachmentsCommand(event) {
  executeWithNotification(
    event,
    removeAttachmentsCore,
    t("commands.notifications.cannotRemoveAttachments")
  );
}

function keepTwoRepliesCommand(event) {
  executeWithNotification(event, keepTwoRepliesCore, t("commands.notifications.cannotKeepReplies"));
}

function cleanAllCommand(event) {
  executeWithNotification(event, cleanAllCore, t("commands.notifications.cannotCleanAll"));
}

function partialReplyCommand(event) {
  executeWithNotification(event, partialReplyCore, t("commands.notifications.cannotPartialReply"));
}

function partialReplyAllCommand(event) {
  executeWithNotification(
    event,
    partialReplyAllCore,
    t("commands.notifications.cannotPartialReplyAll")
  );
}

function partialForwardCommand(event) {
  executeWithNotification(
    event,
    partialForwardCore,
    t("commands.notifications.cannotPartialForward")
  );
}

// Register the function with Office.
Office.actions.associate("removeImagesCommand", removeImagesCommand);
Office.actions.associate("removeAttachmentsCommand", removeAttachmentsCommand);
Office.actions.associate("keepTwoRepliesCommand", keepTwoRepliesCommand);
Office.actions.associate("cleanAllCommand", cleanAllCommand);
Office.actions.associate("partialReplyCommand", partialReplyCommand);
Office.actions.associate("partialReplyAllCommand", partialReplyAllCommand);
Office.actions.associate("partialForwardCommand", partialForwardCommand);
