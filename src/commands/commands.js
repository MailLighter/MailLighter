/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

import { t } from "../shared/i18n";
import {
  sanitizeHtml,
  toHtmlFromText,
  displayFormWithFallback,
  getForwardSubject,
} from "../shared/office-helpers";

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

      reject(new Error(result.error && result.error.message ? result.error.message : t(failedKey)));
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
    const details = error instanceof Error && error.message ? ` (${error.message})` : "";
    notify(`${errorMessage}${details}`);
  } finally {
    event.completed();
  }
}

const MAX_REPLY_BODY_LENGTH = 30000;


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
    const sanitizedSelection = sanitizeHtml(htmlSelection);

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


function openReplyFormAsync(replyAll, htmlBody) {
  const item = Office.context.mailbox.item;
  const errorKey = replyAll
    ? "commands.errors.replyAllFormUnavailable"
    : "commands.errors.replyFormUnavailable";

  if (!item) {
    return Promise.reject(new Error(t(errorKey)));
  }

  const displayFn = replyAll ? item.displayReplyAllForm : item.displayReplyForm;

  if (typeof displayFn !== "function") {
    return Promise.reject(new Error(t(errorKey)));
  }

  try {
    displayFormWithFallback(displayFn.bind(item), htmlBody);
    return Promise.resolve();
  } catch (error) {
    return Promise.reject(error instanceof Error ? error : new Error(String(error)));
  }
}


function openForwardFormAsync(htmlBody) {
  const mailbox = Office.context.mailbox;
  const item = mailbox ? mailbox.item : null;

  try {
    if (item && typeof item.displayForwardForm === "function") {
      displayFormWithFallback(item.displayForwardForm.bind(item), htmlBody);
      return Promise.resolve();
    }

    if (!mailbox || typeof mailbox.displayNewMessageForm !== "function") {
      return Promise.reject(new Error(t("commands.errors.forwardFormUnavailable")));
    }

    const formPayload = {
      subject: getForwardSubject(item && item.subject ? item.subject : ""),
    };

    if (htmlBody) {
      formPayload.htmlBody = htmlBody;
    }

    mailbox.displayNewMessageForm(formPayload);
    return Promise.resolve();
  } catch (error) {
    return Promise.reject(error instanceof Error ? error : new Error(String(error)));
  }
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
  const allAttachments = await getAttachmentsAsync();
  // Keep only real file attachments, not inline images (signatures, logos, etc.)
  const attachments = allAttachments.filter((a) => !a.isInline);

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

function collectRegexPositions(htmlBody, regex, headerCheck) {
  const positions = [];
  let match;
  while ((match = regex.exec(htmlBody)) !== null) {
    if (headerCheck) {
      const after = htmlBody.substring(match.index, match.index + 500);
      if (!headerCheck.test(after)) continue;
    }
    positions.push(match.index);
  }
  return positions;
}

function findTextSeparators(htmlBody) {
  // Search for "De :" / "From :" confirmed by a second standard email header keyword
  // (Envoyé/Sent, Objet/Subject, etc.) within 1500 chars.
  // Allows HTML tags and entities between keyword and colon to handle all Outlook variants
  // (e.g. <b>De</b>&nbsp;: or Envoy&eacute; :).
  const TAG_OR_GAP = "(?:\\s|<[^>]*>|&\\w+;|&#\\d+;|\\xA0)*";
  const fromRegex = new RegExp(
    "\\b(De|From|Von|Van|Da|Fra)" + TAG_OR_GAP + ":",
    "gi"
  );
  // Confirm with any standard reply header field: Envoyé/Sent, Objet/Subject, À/To, Cc.
  // "Objet" and "Subject" have no accents so they survive any HTML encoding.
  const confirmRegex = new RegExp(
    "\\b(Sent|Envoy(?:é|&eacute;|&#233;|e)|Gesendet|Verzonden|Inviato" +
      "|Objet|Subject|Betreff|Onderwerp|Oggetto)" +
      TAG_OR_GAP +
      ":",
    "i"
  );

  const positions = [];
  let match;
  while ((match = fromRegex.exec(htmlBody)) !== null) {
    const after = htmlBody.substring(match.index, match.index + 1500);
    if (!confirmRegex.test(after)) continue;
    // Walk back to the nearest block-level opening tag for a clean cut point
    const lookback = htmlBody.substring(Math.max(0, match.index - 500), match.index);
    const blockTag = lookback.match(/.*(<(?:p|div|tr|li)\b[^>]*>)/is);
    const cutPos = blockTag
      ? match.index - lookback.length + lookback.lastIndexOf(blockTag[1])
      : match.index;
    // Deduplicate: skip if very close to previous position
    if (positions.length > 0 && cutPos - positions[positions.length - 1] < 200) continue;
    positions.push(cutPos);
  }
  return positions;
}

function findReplySeparators(htmlBody) {
  const headerPattern = /\b(From|De|Von|Da|Van|Fra)\s*(&nbsp;|\xA0)?\s*:/i;

  // Strategy 1: divRplyFwdMsg ID (New Outlook, OWA)
  const divPositions = collectRegexPositions(
    htmlBody,
    /<div[^>]*\bid\s*=\s*["'](?:x_)*divRplyFwdMsg["'][^>]*>/gi
  );

  // Strategy 2: border-top styled div (Desktop Outlook / Word)
  const borderPositions = collectRegexPositions(
    htmlBody,
    /<div[^>]*border-top\s*:\s*solid\s[^>]*>/gi,
    headerPattern
  );

  // Strategy 3: <hr> followed by reply header
  const hrPositions = collectRegexPositions(htmlBody, /<hr[^>]*>/gi, headerPattern);

  // Strategy 4: Text-based De:/From: + Objet:/Subject: (handles all HTML encodings)
  const textPositions = findTextSeparators(htmlBody);

  // Return whichever strategy found the most separators
  let best = divPositions;
  if (borderPositions.length > best.length) best = borderPositions;
  if (hrPositions.length > best.length) best = hrPositions;
  if (textPositions.length > best.length) best = textPositions;
  return best;
}

async function keepTwoRepliesCore() {
  const htmlBody = await getBodyAsync(Office.CoercionType.Html);
  const separators = findReplySeparators(htmlBody);

  if (separators.length === 0) {
    return t("commands.notifications.repliesNone");
  }

  if (separators.length <= 2) {
    return t("commands.notifications.repliesNoChange", { count: separators.length });
  }

  // Cut just before the 3rd reply separator.
  // Also pull the cut point back past any visual separator (hr or underscore line)
  // that immediately precedes the 3rd header, so it gets removed too.
  let cutPoint = separators[2];
  const before = htmlBody.substring(0, cutPoint);

  const lastHr = before.lastIndexOf("<hr");
  if (lastHr >= 0 && cutPoint - lastHr < 500) {
    cutPoint = lastHr;
  } else {
    // Look for the last underscore separator block before cutPoint
    let lastUnderscoreIdx = -1;
    const underscoreRe = /_{10,}/g;
    let m;
    while ((m = underscoreRe.exec(before)) !== null) {
      lastUnderscoreIdx = m.index;
    }
    if (lastUnderscoreIdx >= 0 && cutPoint - lastUnderscoreIdx < 2000) {
      const tagStart = before.lastIndexOf("<", lastUnderscoreIdx);
      cutPoint = tagStart >= 0 ? tagStart : lastUnderscoreIdx;
    }
  }

  await setBodyAsync(htmlBody.substring(0, cutPoint), Office.CoercionType.Html);
  return t("commands.notifications.repliesCleaned", { count: separators.length });
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

// Register all commands with Office.
[
  ["removeImagesCommand", removeImagesCore, "cannotRemoveImages"],
  ["removeAttachmentsCommand", removeAttachmentsCore, "cannotRemoveAttachments"],
  ["keepTwoRepliesCommand", keepTwoRepliesCore, "cannotKeepReplies"],
  ["cleanAllCommand", cleanAllCore, "cannotCleanAll"],
  ["partialReplyCommand", partialReplyCore, "cannotPartialReply"],
  ["partialReplyAllCommand", partialReplyAllCore, "cannotPartialReplyAll"],
  ["partialForwardCommand", partialForwardCore, "cannotPartialForward"],
].forEach(([name, core, errorKey]) => {
  Office.actions.associate(name, (event) => {
    executeWithNotification(event, core, t(`commands.notifications.${errorKey}`));
  });
});
