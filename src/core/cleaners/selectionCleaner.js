import { addPendingEvent } from "../savings/pendingSavings";
import { createCleanupEvent } from "../types/CleanupEvent";
import { createCleanupResult } from "../types/CleanupResult";
import { ELEMENT_TYPES } from "../../config/constants";
import { logger } from "../../utils/logger";

/**
 * Replace the draft body with the user's selection only.
 *
 * Preserves the user's composing zone (typing space, append-on-send, signature)
 * by keeping everything before the _MailOriginal anchor / divRplyFwdMsg block,
 * then injecting a thin separator and the sanitized selection.
 */
export async function keepSelectionOnly(platform) {
  // 1. Capture selection BEFORE reading the full body so the cursor context
  //    is still honored.
  const selectedHtml = await platform.getSelectedHtml();

  // 2. Read the full body.
  const fullBody = await platform.getBodyHtml();

  // 3. Find where the quoted/forwarded content starts.
  const mailOriginalMatch = fullBody.match(/<a[^>]*name\s*=\s*["']_MailOriginal["'][^>]*>/i);
  const divRplyMatch = fullBody.match(/<div[^>]*\bid\s*=\s*["']divRplyFwdMsg["'][^>]*>/i);

  let cutPoint = -1;
  if (mailOriginalMatch) cutPoint = mailOriginalMatch.index;
  if (divRplyMatch && (cutPoint === -1 || divRplyMatch.index < cutPoint)) {
    cutPoint = divRplyMatch.index;
  }

  const separator = '<hr style="border:none;border-top:1px solid #b5b5b5;margin:8px 0;">';
  const newBodyHtml =
    cutPoint > 0
      ? fullBody.substring(0, cutPoint) + separator + selectedHtml
      : "<div><br></div>" + separator + selectedHtml;

  const savedBytes = Math.max(0, fullBody.length - newBodyHtml.length);
  await platform.setBodyHtml(newBodyHtml);

  // 4. Place the cursor at the very top by prepending an empty line.
  //    Works in Old Outlook. New Outlook keeps the cursor at the bottom
  //    (known Office.js API limitation). Best-effort only.
  try {
    await platform.prependBodyHtml("<div><br></div>");
  } catch (e) {
    logger.warn("prependBodyHtml unavailable, cursor will stay at bottom", e && e.message);
  }

  if (savedBytes > 0) {
    try {
      const composeId = await platform.getComposeId();
      const recipients = await platform.getRecipients();
      addPendingEvent(
        composeId,
        createCleanupEvent({
          elementType: ELEMENT_TYPES.SELECTION,
          bytesRemoved: savedBytes,
          recipientCount: Array.isArray(recipients) ? recipients.length : 0,
          platform: platform.platformName,
        })
      );
    } catch (e) {
      logger.warn("selectionCleaner pending event failed (non-fatal):", e && e.message);
    }
  }

  return createCleanupResult({
    elementType: ELEMENT_TYPES.SELECTION,
    itemsRemoved: savedBytes > 0 ? 1 : 0,
    bytesRemoved: savedBytes,
  });
}
