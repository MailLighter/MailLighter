/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, console */

import { t } from "../shared/i18n";
import { sanitizeSelectionHtml, toHtmlFromText } from "../shared/office-helpers";

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

function notify(message, icon = "Icon.80x80") {
  const item = Office.context.mailbox.item;

  if (!item || !item.notificationMessages) {
    return;
  }

  if (typeof item.notificationMessages.removeAsync === "function") {
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
    const parts = ['<div><br></div>', separator, selectedHtml];
    await setBodyAsync(parts.join(""), Office.CoercionType.Html);
  }

  // 4. Prepend an empty line to place the cursor at the very top.
  //    Works in Old Outlook. In New Outlook the cursor stays at the bottom
  //    (known Office.js API limitation).
  try {
    await prependAsync('<div><br></div>', Office.CoercionType.Html);
  } catch {
    // prependAsync not available in this client
  }

  return t("commands.notifications.keepSelectionDone");
}

async function removeImagesWork() {
  const htmlBody = await getBodyAsync(Office.CoercionType.Html);
  const imgMatches = htmlBody.match(/<img[^>]*>/gi) || [];

  if (imgMatches.length === 0) {
    return { count: 0, sizeText: "" };
  }

  const cleanedHtml = htmlBody.replace(/<img[^>]*>/gi, "");
  await setBodyAsync(cleanedHtml, Office.CoercionType.Html);

  const totalSize = calculateImageSize(imgMatches);
  return { count: imgMatches.length, sizeText: formatFileSize(totalSize) };
}

async function removeImagesCore() {
  const { count, sizeText } = await removeImagesWork();

  if (count === 0) {
    return t("commands.notifications.imagesNone");
  }

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

  return { count: attachments.length, sizeText: formatFileSize(totalSize) };
}

async function removeAttachmentsCore() {
  const { count, sizeText } = await removeAttachmentsWork();

  if (count === 0) {
    return t("commands.notifications.attachmentsNone");
  }

  return sizeText
    ? t("commands.notifications.attachmentsRemovedWithSize", { count, size: sizeText })
    : t("commands.notifications.attachmentsRemoved", { count });
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
  const TAG_OR_GAP = "(?:\\s|<[^>]*>|&\\w+;|&#\\d+;|\\xA0)*";
  const fromRegex = new RegExp(
    "\\b(De|From|Von|Van|Da|Fra)" + TAG_OR_GAP + ":",
    "gi"
  );
  const confirmRegex = new RegExp(
    "\\b(Sent|Envoy(?:é|&eacute;|&#233;|e)|Enviado|Gesendet|Verzonden|Inviato" +
      "|Objet|Subject|Asunto|Betreff|Onderwerp|Oggetto)" +
      TAG_OR_GAP +
      ":",
    "i"
  );

  const positions = [];
  let match;
  while ((match = fromRegex.exec(htmlBody)) !== null) {
    const after = htmlBody.substring(match.index, match.index + 1500);
    if (!confirmRegex.test(after)) continue;
    const lookback = htmlBody.substring(Math.max(0, match.index - 500), match.index);
    const blockTag = lookback.match(/.*(<(?:p|div|tr|li)\b[^>]*>)/is);
    const cutPos = blockTag
      ? match.index - lookback.length + lookback.lastIndexOf(blockTag[1])
      : match.index;
    if (positions.length > 0 && cutPos - positions[positions.length - 1] < 200) continue;
    positions.push(cutPos);
  }
  return positions;
}

function findReplySeparators(htmlBody) {
  const headerPattern = /\b(From|De|Von|Da|Van|Fra)\s*(&nbsp;|\xA0)?\s*:/i;

  const divPositions = collectRegexPositions(
    htmlBody,
    /<div[^>]*\bid\s*=\s*["'](?:x_)*divRplyFwdMsg["'][^>]*>/gi
  );

  const borderPositions = collectRegexPositions(
    htmlBody,
    /<div[^>]*border-top\s*:\s*solid\s[^>]*>/gi,
    headerPattern
  );

  const hrPositions = collectRegexPositions(htmlBody, /<hr[^>]*>/gi, headerPattern);

  const textPositions = findTextSeparators(htmlBody);

  // Detect Gmail/Apple Mail/Thunderbird inline attributions:
  // "... a écrit :", "... wrote:", "... escribió:", "... schrieb ...:"
  const wroteRegex = /\b(a\s+[eé]crit|wrote|escribi[oó]|escribe|schrieb|geschreven|scrisse)\s*:/gi;
  const wrotePositions = [];
  let wroteMatch;
  while ((wroteMatch = wroteRegex.exec(htmlBody)) !== null) {
    const lookback = htmlBody.substring(Math.max(0, wroteMatch.index - 500), wroteMatch.index);
    const blockTag = lookback.match(/.*(<(?:p|div|blockquote|li)\b[^>]*>)/is);
    const cutPos = blockTag
      ? wroteMatch.index - lookback.length + lookback.lastIndexOf(blockTag[1])
      : wroteMatch.index;
    if (wrotePositions.length > 0 && cutPos - wrotePositions[wrotePositions.length - 1] < 200)
      continue;
    wrotePositions.push(cutPos);
  }

  let best = divPositions;
  if (borderPositions.length > best.length) best = borderPositions;
  if (hrPositions.length > best.length) best = hrPositions;
  if (textPositions.length > best.length) best = textPositions;
  if (wrotePositions.length > best.length) best = wrotePositions;
  return best;
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
  const before = htmlBody.substring(0, cutPoint);

  const lastHr = before.lastIndexOf("<hr");
  if (lastHr >= 0 && cutPoint - lastHr < 500) {
    cutPoint = lastHr;
  } else {
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
  return { found: separators.length, cleaned: true };
}

async function keepTwoRepliesCore() {
  const { found, cleaned } = await keepTwoRepliesWork();

  if (found === 0) {
    return t("commands.notifications.repliesNone");
  }

  if (!cleaned) {
    return t("commands.notifications.repliesNoChange", { count: found });
  }

  return t("commands.notifications.repliesCleaned", { count: found });
}

function formatCleanAllPart(prefix, count, sizeText) {
  if (!count) return prefix + ": 0";
  return sizeText ? `${prefix}: ${count} (${sizeText})` : `${prefix}: ${count}`;
}

async function cleanAllCore() {
  const parts = [];

  try {
    const img = await removeImagesWork();
    parts.push(formatCleanAllPart(t("commands.notifications.cleanAllImagesPrefix"), img.count, img.sizeText));
  } catch (error) {
    parts.push(`${t("commands.notifications.cleanAllImagesPrefix")}: ${error.message}`);
  }

  try {
    const att = await removeAttachmentsWork();
    parts.push(formatCleanAllPart(t("commands.notifications.cleanAllAttachmentsPrefix"), att.count, att.sizeText));
  } catch (error) {
    parts.push(`${t("commands.notifications.cleanAllAttachmentsPrefix")}: ${error.message}`);
  }

  try {
    const rep = await keepTwoRepliesWork();
    const prefix = t("commands.notifications.cleanAllRepliesPrefix");
    if (rep.found === 0) {
      parts.push(prefix + ": 0");
    } else if (rep.cleaned) {
      parts.push(`${prefix}: ${rep.found} → 2`);
    } else {
      parts.push(`${prefix}: ${rep.found}`);
    }
  } catch (error) {
    parts.push(`${t("commands.notifications.cleanAllRepliesPrefix")}: ${error.message}`);
  }

  return t("commands.notifications.cleanAllDone", { details: parts.join(" | ") });
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
