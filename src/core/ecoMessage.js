import { STORAGE_KEYS, MAILLIGHTER_SITE_URL } from "../config/constants";
import { escapeHtml } from "../utils/format";

/**
 * Read whether the user enabled the ecological message footer.
 */
export function isEcoMessageEnabled(storage) {
  return storage.getItem(STORAGE_KEYS.ECO_MESSAGE_ENABLED) === "1";
}

/**
 * Read the user's customized ecological message text, falling back to the
 * provided default (which is i18n-bound and supplied by the router/UI).
 */
export function getEcoMessageText(storage, defaultText) {
  return storage.getItem(STORAGE_KEYS.ECO_MESSAGE_TEXT) || defaultText;
}

/**
 * Append a styled "ecological message" footer to the current draft body.
 * The text is HTML-escaped and the literal "MailLighter" word is linked back
 * to the public site (constant URL — never user-controlled).
 */
export async function appendEcoMessage(platform, text) {
  const htmlBody = await platform.getBodyHtml();
  const safeText = escapeHtml(text);
  const linkedText = safeText.replace(
    /MailLighter/g,
    `<a href="${MAILLIGHTER_SITE_URL}" style="color:#1b5e20;">MailLighter</a>`
  );
  const ecoHtml =
    `<div style="margin-top:12px;padding-top:8px;border-top:1px solid #c8e6c9;` +
    `color:#2e7d32;font-size:13px;">${linkedText}</div>`;
  await platform.setBodyHtml(htmlBody + ecoHtml);
}
