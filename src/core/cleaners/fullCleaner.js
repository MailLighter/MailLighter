import { removeImages } from "./imageCleaner";
import { removeAttachments } from "./attachmentCleaner";
import { keepTwoReplies } from "./replyCleaner";

/**
 * Run all cleanup actions in sequence. Each step is independently
 * fault-tolerant: a failure in one cleaner does not abort the others.
 *
 * Returns { images, attachments, replies, errors } where each cleanup result
 * is either a CleanupResult or null (if it failed), and errors carries the
 * thrown error keys for router-side notification.
 */
export async function cleanAll(platform) {
  const out = {
    images: null,
    attachments: null,
    replies: null,
    errors: { images: null, attachments: null, replies: null },
  };

  try {
    out.images = await removeImages(platform);
  } catch (e) {
    out.errors.images = e;
  }

  try {
    out.attachments = await removeAttachments(platform);
  } catch (e) {
    out.errors.attachments = e;
  }

  try {
    out.replies = await keepTwoReplies(platform);
  } catch (e) {
    out.errors.replies = e;
  }

  return out;
}
