/* global localStorage */

/**
 * pendingSavings — local queue of "candidate" savings, confirmed (and
 * recorded into the Settings counter) ONLY when the user actually sends
 * the email.
 *
 * Lifecycle:
 *   1. cleaner calls addPendingEvent(composeId, event)
 *   2a. user clicks "Send" → sendHandler calls consumePendingEvents
 *       → events confirmed via recordConfirmedSavings → queue cleared
 *   2b. user closes without sending → entry expires via TTL after 24h
 *
 * Storage: localStorage when available (persists across JS runtimes in
 * Outlook — critical for OnMessageSend which runs in a separate context
 * from the ribbon command handlers). Falls back to an in-memory Map in
 * non-browser environments (Node.js / Jest).
 */

import { logger } from "../../utils/logger";

const PENDING_PREFIX = "maillighter_pending_";
const MAX_AGE_MS = 24 * 60 * 60 * 1000;

// In-memory fallback for non-browser environments (tests / Node.js).
const memFallback = new Map();

// Secondary in-memory index used by getPendingState() and purgeStale().
// Accurate within a single JS context; not needed to be cross-runtime.
const activeIds = new Set();

function _ls() {
  return typeof localStorage !== "undefined" ? localStorage : null;
}

function _readEntry(composeId) {
  const ls = _ls();
  if (ls) {
    try {
      const raw = ls.getItem(PENDING_PREFIX + composeId);
      return raw ? JSON.parse(raw) : null;
    } catch {
      return null;
    }
  }
  return memFallback.get(composeId) || null;
}

function _writeEntry(composeId, entry) {
  const ls = _ls();
  if (ls) {
    try {
      ls.setItem(PENDING_PREFIX + composeId, JSON.stringify(entry));
    } catch (e) {
      logger.warn("pendingSavings: localStorage write failed", e && e.message);
    }
  } else {
    memFallback.set(composeId, entry);
  }
}

function _deleteEntry(composeId) {
  const ls = _ls();
  if (ls) {
    ls.removeItem(PENDING_PREFIX + composeId);
  } else {
    memFallback.delete(composeId);
  }
}

/**
 * Records a candidate saving linked to the in-progress compose.
 * @param {string} composeId
 * @param {CleanupEvent} event
 */
export function addPendingEvent(composeId, event) {
  if (!composeId) {
    logger.warn("addPendingEvent called without composeId, event dropped");
    return;
  }
  const existing = _readEntry(composeId) || { events: [], createdAt: Date.now() };
  existing.events.push(event);
  _writeEntry(composeId, existing);
  activeIds.add(composeId);
}

/**
 * Retrieves and CLEARS all candidate savings for a confirmed send.
 * Called by sendHandler.
 * @param {string} composeId
 * @returns {Array<CleanupEvent>}
 */
export function consumePendingEvents(composeId) {
  const entry = _readEntry(composeId);
  _deleteEntry(composeId);
  activeIds.delete(composeId);
  return entry ? entry.events : [];
}

/**
 * Clears candidate savings without confirming (rare; usually entries expire
 * via TTL purgeStale instead).
 * @param {string} composeId
 */
export function discardPendingEvents(composeId) {
  _deleteEntry(composeId);
  activeIds.delete(composeId);
}

/**
 * Inspects pending state without mutating (for tests/debug).
 * @returns {{ count: number }}
 */
export function getPendingState() {
  return { count: activeIds.size };
}

/**
 * Purges entries older than maxAgeMs. Should be invoked periodically to
 * avoid a memory leak when a compose window stays open without ever
 * sending or closing.
 */
export function purgeStale(maxAgeMs = MAX_AGE_MS) {
  const now = Date.now();
  const ls = _ls();
  if (ls) {
    for (const composeId of [...activeIds]) {
      const entry = _readEntry(composeId);
      if (!entry || now - entry.createdAt > maxAgeMs) {
        _deleteEntry(composeId);
        activeIds.delete(composeId);
      }
    }
  } else {
    for (const [composeId, entry] of memFallback.entries()) {
      if (now - entry.createdAt > maxAgeMs) {
        memFallback.delete(composeId);
        activeIds.delete(composeId);
      }
    }
  }
}

/**
 * Test-only: resets all pending state. Not exposed in production paths.
 */
export function _resetPendingForTests() {
  const ls = _ls();
  if (ls) {
    for (const composeId of [...activeIds]) {
      ls.removeItem(PENDING_PREFIX + composeId);
    }
  }
  memFallback.clear();
  activeIds.clear();
}
