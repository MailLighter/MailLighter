/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, window, URLSearchParams, globalThis */

import { t } from "../ui/i18n";
import { formatFileSize } from "../utils/format";
import { logger } from "../utils/logger";
import { STORAGE_KEYS } from "../config/constants";
import { OutlookAdapter } from "../platforms/outlook/outlookAdapter";
import { createStorage, getSavings } from "../core/savings/savingsCalculator";
import { onMessageSend, confirmEagerly } from "../core/lifecycle/sendHandler";
import { removeImages } from "../core/cleaners/imageCleaner";
import { removeAttachments } from "../core/cleaners/attachmentCleaner";
import { keepTwoReplies } from "../core/cleaners/replyCleaner";
import { keepSelectionOnly } from "../core/cleaners/selectionCleaner";
import { cleanAll } from "../core/cleaners/fullCleaner";
import { isEcoMessageEnabled, getEcoMessageText, appendEcoMessage } from "../core/ecoMessage";

Office.onReady(() => {});

const platform = new OutlookAdapter();
const storage = createStorage();

function unitLabels() {
  return {
    kilobytes: t("units.kilobytes"),
    megabytes: t("units.megabytes"),
    gigabytes: t("units.gigabytes"),
    lessThanOne: t("units.lessThanOne"),
  };
}

function formatBytes(bytes) {
  return formatFileSize(bytes, { units: unitLabels() });
}

function currentEcoText() {
  return getEcoMessageText(storage, t("settings.ecoMessageDefault"));
}

/**
 * Translates an error key thrown by the adapter/cleaners into a user-facing
 * notification text, falling back to the generic key when the thrown message
 * is not a known i18n path.
 */
function translateError(error, fallbackKey) {
  if (error instanceof Error && error.message) {
    const translated = t(error.message);
    if (translated !== error.message) return translated;
    logger.warn(error.message);
  }
  return t(fallbackKey);
}

async function executeWithNotification(event, worker, errorKey) {
  try {
    const successMessage = await worker();
    platform.notify(successMessage);
    if (!platform.supportsOnMessageSend()) {
      try {
        const composeId = await platform.getComposeId();
        await confirmEagerly(composeId, storage);
      } catch (eagerError) {
        logger.warn("confirmEagerly failed (non-fatal):", eagerError && eagerError.message);
      }
    }
  } catch (error) {
    platform.notify(translateError(error, errorKey));
  } finally {
    event.completed();
  }
}

async function removeImagesCommand(event) {
  await executeWithNotification(
    event,
    async () => {
      const result = await removeImages(platform);
      if (result.itemsRemoved === 0) {
        return t("commands.notifications.imagesNone");
      }
      const sizeText = result.bytesRemoved > 0 ? formatBytes(result.bytesRemoved) : "";
      return sizeText
        ? t("commands.notifications.imagesRemovedWithSize", {
            count: result.itemsRemoved,
            size: sizeText,
          })
        : t("commands.notifications.imagesRemoved", { count: result.itemsRemoved });
    },
    "commands.notifications.cannotRemoveImages"
  );
}

async function removeAttachmentsCommand(event) {
  await executeWithNotification(
    event,
    async () => {
      const result = await removeAttachments(platform);
      if (result.itemsRemoved === 0) {
        return t("commands.notifications.attachmentsNone");
      }
      const sizeText = result.bytesRemoved > 0 ? formatBytes(result.bytesRemoved) : "";
      return sizeText
        ? t("commands.notifications.attachmentsRemovedWithSize", {
            count: result.itemsRemoved,
            size: sizeText,
          })
        : t("commands.notifications.attachmentsRemoved", { count: result.itemsRemoved });
    },
    "commands.notifications.cannotRemoveAttachments"
  );
}

async function keepTwoRepliesCommand(event) {
  await executeWithNotification(
    event,
    async () => {
      const result = await keepTwoReplies(platform);
      const found = typeof result.found === "number" ? result.found : 0;

      if (found === 0) {
        return t("commands.notifications.repliesNone");
      }
      if (result.bytesRemoved === 0) {
        return t("commands.notifications.repliesNoChange", { count: found });
      }

      if (isEcoMessageEnabled(storage)) {
        try {
          await appendEcoMessage(platform, currentEcoText());
        } catch (ecoError) {
          logger.warn("appendEcoMessage failed (non-fatal):", ecoError && ecoError.message);
        }
      }

      const sizeText = formatBytes(result.bytesRemoved);
      return sizeText
        ? t("commands.notifications.repliesCleanedWithSize", { count: found, size: sizeText })
        : t("commands.notifications.repliesCleaned", { count: found });
    },
    "commands.notifications.cannotKeepReplies"
  );
}

async function keepSelectionOnlyCommand(event) {
  await executeWithNotification(
    event,
    async () => {
      const result = await keepSelectionOnly(platform);
      const sizeText = result.bytesRemoved > 0 ? formatBytes(result.bytesRemoved) : "";
      return sizeText
        ? t("commands.notifications.keepSelectionDoneWithSize", { size: sizeText })
        : t("commands.notifications.keepSelectionDone");
    },
    "commands.notifications.cannotKeepSelection"
  );
}

function formatCleanAllPart(prefix, count, sizeText) {
  const colon = t("units.colon");
  if (!count) return prefix + colon + "0";
  return sizeText ? `${prefix}${colon}${count} (${sizeText})` : `${prefix}${colon}${count}`;
}

async function cleanAllCommand(event) {
  await executeWithNotification(
    event,
    async () => {
      const out = await cleanAll(platform);
      const parts = [];
      let totalBytes = 0;

      // Images
      if (out.errors.images) {
        parts.push(
          `${t("commands.notifications.cleanAllImagesPrefix")}: ${
            out.errors.images instanceof Error ? out.errors.images.message : ""
          }`
        );
      } else {
        const r = out.images;
        parts.push(
          formatCleanAllPart(
            t("commands.notifications.cleanAllImagesPrefix"),
            r.itemsRemoved,
            r.bytesRemoved > 0 ? formatBytes(r.bytesRemoved) : ""
          )
        );
        totalBytes += r.bytesRemoved || 0;
      }

      // Attachments
      if (out.errors.attachments) {
        parts.push(
          `${t("commands.notifications.cleanAllAttachmentsPrefix")}: ${
            out.errors.attachments instanceof Error ? out.errors.attachments.message : ""
          }`
        );
      } else {
        const r = out.attachments;
        parts.push(
          formatCleanAllPart(
            t("commands.notifications.cleanAllAttachmentsPrefix"),
            r.itemsRemoved,
            r.bytesRemoved > 0 ? formatBytes(r.bytesRemoved) : ""
          )
        );
        totalBytes += r.bytesRemoved || 0;
      }

      // Replies
      if (out.errors.replies) {
        parts.push(
          `${t("commands.notifications.cleanAllRepliesPrefix")}: ${
            out.errors.replies instanceof Error ? out.errors.replies.message : ""
          }`
        );
      } else {
        const r = out.replies;
        const found = typeof r.found === "number" ? r.found : 0;
        const prefix = t("commands.notifications.cleanAllRepliesPrefix");
        const colon = t("units.colon");
        if (found === 0) {
          parts.push(prefix + colon + "0");
        } else if (r.bytesRemoved > 0) {
          const repSizeText = formatBytes(r.bytesRemoved);
          parts.push(
            repSizeText
              ? `${prefix}${colon}${found} → 2 (${repSizeText})`
              : `${prefix}${colon}${found} → 2`
          );
          totalBytes += r.bytesRemoved || 0;
          if (isEcoMessageEnabled(storage)) {
            try {
              await appendEcoMessage(platform, currentEcoText());
            } catch (ecoError) {
              logger.warn("appendEcoMessage failed (non-fatal):", ecoError && ecoError.message);
            }
          }
        } else {
          parts.push(`${prefix}${colon}${found}`);
        }
      }

      const totalText =
        totalBytes > 0
          ? t("commands.notifications.cleanAllTotal", { size: formatBytes(totalBytes) })
          : "";

      return t("commands.notifications.cleanAllDone", {
        details: parts.join(" | "),
        total: totalText,
      });
    },
    "commands.notifications.cannotCleanAll"
  );
}

function openSettingsCommand(event) {
  const ecoEnabled = isEcoMessageEnabled(storage);
  const savings = getSavings(storage);
  const params = new URLSearchParams({
    ecoMessage: ecoEnabled ? "1" : "0",
    ecoText: currentEcoText(),
    transImages: String(savings.transmission.images),
    transReplies: String(savings.transmission.replies),
    transAttachments: String(savings.transmission.attachments),
    transTotal: String(savings.transmission.total),
  });
  const settingsUrl = new window.URL(`settings.html?${params}`, window.location.href).toString();

  Office.context.ui.displayDialogAsync(settingsUrl, { height: 65, width: 40 }, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      platform.notify(t("commands.notifications.cannotOpenSettings"));
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
        storage.setItem(STORAGE_KEYS.ECO_MESSAGE_ENABLED, data.ecoMessageEnabled ? "1" : "0");
      }
      if (typeof data.ecoMessageText !== "undefined") {
        storage.setItem(STORAGE_KEYS.ECO_MESSAGE_TEXT, data.ecoMessageText);
      }
      if (data.action === "close") {
        try {
          dialog.close();
        } catch {
          // dialog handle stale — dialog closed itself via window.close()
        }
        finish();
      }
    });

    dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
      finish();
    });
  });
}

// Register all commands with Office.
Office.actions.associate("removeImagesCommand", removeImagesCommand);
Office.actions.associate("removeAttachmentsCommand", removeAttachmentsCommand);
Office.actions.associate("keepTwoRepliesCommand", keepTwoRepliesCommand);
Office.actions.associate("cleanAllCommand", cleanAllCommand);
Office.actions.associate("keepSelectionOnlyCommand", keepSelectionOnlyCommand);
Office.actions.associate("openSettingsCommand", openSettingsCommand);

// Expose the OnMessageSend handler globally so the LaunchEvent declared in
// manifest.xml can find it. webpack must not tree-shake this assignment.
const onMessageSendGlobalHandler = (eventArgs) => onMessageSend(eventArgs, platform, storage);
if (typeof globalThis !== "undefined") {
  globalThis.onMessageSendGlobalHandler = onMessageSendGlobalHandler;
}

// Re-export to anchor the symbol against tree-shaking.
export { onMessageSendGlobalHandler };

// Also bind via Office.actions.associate as a fallback some Outlook clients
// need for LaunchEvent dispatch.
if (Office.actions && typeof Office.actions.associate === "function") {
  Office.actions.associate("onMessageSendGlobalHandler", onMessageSendGlobalHandler);
}
