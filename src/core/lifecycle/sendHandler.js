/**
 * sendHandler — invoked by Office.js (OnMessageSend launch event) when the
 * user clicks "Send".
 *
 * Phase 1a responsibilities:
 *   1. Resolve the current composeId
 *   2. Pull pending events for that composeId
 *   3. Recompute the CURRENT recipient count (may have changed since cleanup)
 *   4. Confirm each event into the Settings counter via recordConfirmedSavings
 *   5. ALWAYS allow the send to proceed, even on internal failure
 *
 * Phase 1b (future): also call collectEvent (tenant telemetry, never per-user).
 */

import { consumePendingEvents } from "../savings/pendingSavings";
import { recordConfirmedSavings } from "../savings/savingsCalculator";
import { logger } from "../../utils/logger";

/**
 * @param {object} eventArgs - Office.js OnMessageSend event arguments
 * @param {PlatformAdapter} platform
 * @param {object} storage - storage abstraction (createStorage())
 */
export async function onMessageSend(eventArgs, platform, storage) {
  try {
    const composeId = await platform.getComposeId();
    const events = consumePendingEvents(composeId);

    if (events.length === 0) {
      return;
    }

    const currentRecipients = await platform.getRecipients();
    const currentCount = Array.isArray(currentRecipients) ? currentRecipients.length : 0;

    for (const event of events) {
      const updated = { ...event, recipientCount: currentCount };
      recordConfirmedSavings(storage, updated);
    }

    logger.info(`onMessageSend: ${events.length} pending event(s) confirmed`);
  } catch (e) {
    logger.warn("sendHandler error (silently ignored)", e);
  } finally {
    // CRITICAL: never block the send.
    if (eventArgs && typeof eventArgs.completed === "function") {
      eventArgs.completed({ allowEvent: true });
    }
  }
}

/**
 * Confirms pending savings immediately, for legacy clients (Mailbox < 1.10)
 * where OnMessageSend never fires.
 *
 * Uses the recipientCount already stored in each pending event (captured at
 * cleanup time) rather than re-reading recipients — acceptable trade-off since
 * confirmEagerly is called right after the cleanup action.
 *
 * @param {string} composeId
 * @param {object} storage - storage abstraction (createStorage())
 */
export async function confirmEagerly(composeId, storage) {
  try {
    if (!composeId) return;
    const events = consumePendingEvents(composeId);
    if (events.length === 0) return;
    for (const event of events) {
      recordConfirmedSavings(storage, event);
    }
    logger.info(`confirmEagerly: ${events.length} pending event(s) confirmed (legacy path)`);
  } catch (e) {
    logger.warn("confirmEagerly error (silently ignored)", e);
  }
}
