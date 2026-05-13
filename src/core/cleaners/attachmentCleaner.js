import { recordConfirmedSavings } from "../savings/savingsCalculator";
import { createCleanupEvent } from "../types/CleanupEvent";
import { createCleanupResult } from "../types/CleanupResult";
import { ELEMENT_TYPES } from "../../config/constants";
import { logger } from "../../utils/logger";

/**
 * Remove all non-inline file attachments from the current draft.
 * Inline images (signatures, logos) are preserved.
 * The Settings counter is updated immediately, using the recipient count
 * captured at click time.
 */
export async function removeAttachments(platform, storage) {
  const allAttachments = await platform.listAttachments();
  const attachments = (allAttachments || []).filter((a) => !a.isInline);

  if (attachments.length === 0) {
    return createCleanupResult({
      elementType: ELEMENT_TYPES.ATTACHMENT,
      itemsRemoved: 0,
      bytesRemoved: 0,
    });
  }

  const totalSize = attachments.reduce((sum, attachment) => {
    const size = typeof attachment.size === "number" ? attachment.size : 0;
    return sum + size;
  }, 0);

  await Promise.all(attachments.map((attachment) => platform.removeAttachment(attachment.id)));

  if (totalSize > 0 && storage) {
    try {
      const recipients = await platform.getRecipients();
      recordConfirmedSavings(
        storage,
        createCleanupEvent({
          elementType: ELEMENT_TYPES.ATTACHMENT,
          bytesRemoved: totalSize,
          recipientCount: Array.isArray(recipients) ? recipients.length : 0,
          platform: platform.platformName,
        })
      );
    } catch (e) {
      logger.warn("attachmentCleaner savings record failed (non-fatal):", e && e.message);
    }
  }

  return createCleanupResult({
    elementType: ELEMENT_TYPES.ATTACHMENT,
    itemsRemoved: attachments.length,
    bytesRemoved: totalSize,
  });
}
