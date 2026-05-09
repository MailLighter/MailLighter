/* global localStorage */

import { ELEMENT_TYPE_TO_CATEGORY, STORAGE_KEYS } from "../../config/constants";
import { logger } from "../../utils/logger";

const TRANSMISSION_KEYS = {
  images: STORAGE_KEYS.USER_SAVINGS_TRANSMISSION_IMAGES,
  replies: STORAGE_KEYS.USER_SAVINGS_TRANSMISSION_REPLIES,
  attachments: STORAGE_KEYS.USER_SAVINGS_TRANSMISSION_ATTACHMENTS,
};

function readCount(storage, key) {
  const raw = storage.getItem(key);
  const parsed = parseInt(raw || "0", 10);
  return Number.isFinite(parsed) ? parsed : 0;
}

function addToKey(storage, key, bytes) {
  if (!bytes || bytes <= 0) return;
  storage.setItem(key, String(readCount(storage, key) + bytes));
}

/**
 * Default storage abstraction over the browser localStorage. Tests can pass
 * a custom storage with the same interface.
 */
export function createStorage() {
  if (typeof localStorage === "undefined") {
    const mem = new Map();
    return {
      getItem: (k) => (mem.has(k) ? mem.get(k) : null),
      setItem: (k, v) => mem.set(k, String(v)),
    };
  }
  return localStorage;
}

/**
 * Confirms a cleanup event at send time and updates BOTH counters:
 *  - bytesRemoved (raw bytes physically removed from the user's draft)
 *  - bytesRemoved × recipientCount (transmission savings — how many bytes
 *    were avoided across all server hops + recipient mailboxes)
 *
 * Called only by sendHandler. Cleaners must NOT call this.
 *
 * @param {object} storage  - storage abstraction (createStorage() by default)
 * @param {CleanupEvent} event  - event with recipientCount RECALCULATED at send
 */
export function recordConfirmedSavings(storage, event) {
  if (!event || typeof event.bytesRemoved !== "number" || event.bytesRemoved <= 0) {
    return;
  }
  const category = ELEMENT_TYPE_TO_CATEGORY[event.elementType];
  if (!category) {
    logger.warn("recordConfirmedSavings: unknown elementType", event.elementType);
    return;
  }
  const recipients = Math.max(1, event.recipientCount || 0);
  const transmissionBytes = event.bytesRemoved * recipients;

  addToKey(storage, TRANSMISSION_KEYS[category], transmissionBytes);
}

/**
 * Reads cumulative savings for the Settings UI.
 *
 *   raw.{images,replies,attachments,total}        -- bytes physically removed
 *   transmission.{images,replies,attachments,total} -- bytes × recipients
 */
export function getSavings(storage) {
  const transImages = readCount(storage, TRANSMISSION_KEYS.images);
  const transReplies = readCount(storage, TRANSMISSION_KEYS.replies);
  const transAttachments = readCount(storage, TRANSMISSION_KEYS.attachments);

  return {
    transmission: {
      images: transImages,
      replies: transReplies,
      attachments: transAttachments,
      total: transImages + transReplies + transAttachments,
    },
  };
}
