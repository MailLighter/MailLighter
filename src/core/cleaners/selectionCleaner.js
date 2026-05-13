import { recordConfirmedSavings } from "../savings/savingsCalculator";
import { createCleanupEvent } from "../types/CleanupEvent";
import { createCleanupResult } from "../types/CleanupResult";
import { ELEMENT_TYPES } from "../../config/constants";
import { logger } from "../../utils/logger";

/**
 * Replace the draft body with the user's selection only.
 *
 * Preserves the user's composing zone (typing space, append-on-send, signature)
 * by keeping everything before the _MailOriginal anchor / divRplyFwdMsg block,
 * then injecting the sanitized selection above the signature (when present)
 * so the kept text reads as the message body, not as a quote below the sign-off.
 */
export async function keepSelectionOnly(platform, storage) {
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
  const keptPart = cutPoint > 0 ? fullBody.substring(0, cutPoint) : "<div><br></div>";

  // 4. Within the kept part, locate the signature so we can insert the
  //    selection above it. Outlook (Desktop, OWA, New Outlook) wraps the
  //    signature in <div id="Signature"> or id="ms-outlook-mobile-signature">.
  const signatureMatch =
    keptPart.match(/<div[^>]*\bid\s*=\s*["']Signature["'][^>]*>/i) ||
    keptPart.match(/<div[^>]*\bid\s*=\s*["']ms-outlook-mobile-signature["'][^>]*>/i);

  const newBodyHtml = signatureMatch
    ? keptPart.substring(0, signatureMatch.index) +
      selectedHtml +
      separator +
      keptPart.substring(signatureMatch.index)
    : keptPart + separator + selectedHtml;

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

  if (savedBytes > 0 && storage) {
    try {
      const recipients = await platform.getRecipients();
      recordConfirmedSavings(
        storage,
        createCleanupEvent({
          elementType: ELEMENT_TYPES.SELECTION,
          bytesRemoved: savedBytes,
          recipientCount: Array.isArray(recipients) ? recipients.length : 0,
          platform: platform.platformName,
        })
      );
    } catch (e) {
      logger.warn("selectionCleaner savings record failed (non-fatal):", e && e.message);
    }
  }

  return createCleanupResult({
    elementType: ELEMENT_TYPES.SELECTION,
    itemsRemoved: savedBytes > 0 ? 1 : 0,
    bytesRemoved: savedBytes,
  });
}
