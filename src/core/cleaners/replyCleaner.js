import { recordConfirmedSavings } from "../savings/savingsCalculator";
import { createCleanupEvent } from "../types/CleanupEvent";
import { createCleanupResult } from "../types/CleanupResult";
import { ELEMENT_TYPES, CLEANUP_DEFAULTS } from "../../config/constants";
import { findReplySeparators } from "../replyDetection";
import { logger } from "../../utils/logger";

/**
 * Pure logic: given an HTML body, return the truncation point that keeps the
 * latest N replies (default 2). Returns { found, cleaned, cutPoint, savedBytes }.
 *
 * Snap-back rules:
 *   - <hr> immediately preceding the cut → snap to the <hr>
 *   - "-----Original Message-----" style separators → snap to the dash line
 *     (but NOT onto "Forwarded Message" variants — that content is intentional
 *     forwarded metadata the user wants to keep)
 */
export function computeKeepTwoRepliesCut(
  htmlBody,
  keepCount = CLEANUP_DEFAULTS.KEEP_REPLIES_COUNT
) {
  const separators = findReplySeparators(htmlBody);
  if (separators.length === 0) {
    return { found: 0, cleaned: false, cutPoint: -1, savedBytes: 0 };
  }
  if (separators.length <= keepCount) {
    return { found: separators.length, cleaned: false, cutPoint: -1, savedBytes: 0 };
  }

  let cutPoint = separators[keepCount];
  const floor = separators[keepCount - 1];
  const before = htmlBody.substring(0, cutPoint);

  const lastHr = before.lastIndexOf("<hr");
  if (lastHr >= 0 && cutPoint - lastHr < 500 && lastHr >= floor) {
    cutPoint = lastHr;
  } else {
    const dashSepRe = /[-‐-—]{3,}[ \t\xa0]*[^-‐-—\n\r<]{3,60}[ \t\xa0]*[-‐-—]{3,}/g;
    const forwardedLabelRe =
      /Forwarded Message|Message transf(?:é|&eacute;|&#233;)r(?:é|&eacute;|&#233;)|Mensaje reenviado|Weitergeleitete Nachricht|Doorgestuurd bericht|Messaggio inoltrato/i;
    let lastDashSepIdx = -1;
    let dashMatch;
    while ((dashMatch = dashSepRe.exec(before)) !== null) {
      if (dashMatch.index >= floor && !forwardedLabelRe.test(dashMatch[0])) {
        lastDashSepIdx = dashMatch.index;
      }
    }
    if (lastDashSepIdx >= 0 && cutPoint - lastDashSepIdx < 1000) {
      const tagStart = before.lastIndexOf("<", lastDashSepIdx);
      cutPoint = tagStart >= 0 && tagStart >= floor ? tagStart : lastDashSepIdx;
    }
  }

  const savedBytes = htmlBody.length - cutPoint;
  return { found: separators.length, cleaned: true, cutPoint, savedBytes };
}

/**
 * Truncates the draft body to keep only the latest N replies.
 * Returns a CleanupResult and records the saving immediately if bytes
 * were removed.
 */
export async function keepTwoReplies(platform, storage) {
  const html = await platform.getBodyHtml();
  const { found, cleaned, cutPoint, savedBytes } = computeKeepTwoRepliesCut(html);

  if (!cleaned) {
    return {
      ...createCleanupResult({
        elementType: ELEMENT_TYPES.REPLY,
        itemsRemoved: 0,
        bytesRemoved: 0,
      }),
      found,
    };
  }

  await platform.setBodyHtml(html.substring(0, cutPoint));

  if (savedBytes > 0 && storage) {
    try {
      const recipients = await platform.getRecipients();
      recordConfirmedSavings(
        storage,
        createCleanupEvent({
          elementType: ELEMENT_TYPES.REPLY,
          bytesRemoved: savedBytes,
          recipientCount: Array.isArray(recipients) ? recipients.length : 0,
        })
      );
    } catch (e) {
      logger.warn("replyCleaner savings record failed (non-fatal):", e && e.message);
    }
  }

  return {
    ...createCleanupResult({
      elementType: ELEMENT_TYPES.REPLY,
      itemsRemoved: found,
      bytesRemoved: savedBytes,
    }),
    found,
  };
}
