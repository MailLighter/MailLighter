import { addPendingEvent } from "../savings/pendingSavings";
import { createCleanupEvent } from "../types/CleanupEvent";
import { createCleanupResult } from "../types/CleanupResult";
import { ELEMENT_TYPES } from "../../config/constants";
import { logger } from "../../utils/logger";

/**
 * Heuristic image-size estimation from inline <img> tags:
 *   - explicit data-size attribute → trusted
 *   - else width/height dimensions → ~0.5 bit/pixel rounded to 5 KB blocks
 *   - else flat 50 KB fallback
 */
export function calculateImageSize(imgMatches) {
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

export function stripInlineImages(html) {
  const imgMatches = html.match(/<img[^>]*>/gi) || [];
  if (imgMatches.length === 0) {
    return { cleaned: html, bytesRemoved: 0, imagesRemoved: 0 };
  }
  const cleaned = html.replace(/<img[^>]*>/gi, "");
  const bytesRemoved = calculateImageSize(imgMatches);
  return { cleaned, bytesRemoved, imagesRemoved: imgMatches.length };
}

/**
 * Remove all inline <img> tags from the current draft body.
 * Adds a pending event; the Settings counter is updated only at send time.
 */
export async function removeImages(platform) {
  const html = await platform.getBodyHtml();
  const { cleaned, bytesRemoved, imagesRemoved } = stripInlineImages(html);

  if (imagesRemoved === 0) {
    return createCleanupResult({
      elementType: ELEMENT_TYPES.IMAGE,
      itemsRemoved: 0,
      bytesRemoved: 0,
    });
  }

  await platform.setBodyHtml(cleaned);

  if (bytesRemoved > 0) {
    try {
      const composeId = await platform.getComposeId();
      const recipients = await platform.getRecipients();
      addPendingEvent(
        composeId,
        createCleanupEvent({
          elementType: ELEMENT_TYPES.IMAGE,
          bytesRemoved,
          recipientCount: Array.isArray(recipients) ? recipients.length : 0,
          platform: platform.platformName,
        })
      );
    } catch (e) {
      logger.warn("imageCleaner pending event failed (non-fatal):", e && e.message);
    }
  }

  return createCleanupResult({
    elementType: ELEMENT_TYPES.IMAGE,
    itemsRemoved: imagesRemoved,
    bytesRemoved,
  });
}
