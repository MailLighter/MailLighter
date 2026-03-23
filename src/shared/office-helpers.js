/* global Office */

import { t } from "./i18n";

export function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

export function sanitizeHtml(html) {
  return html
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "")
    .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, "")
    .replace(/<!--[\s\S]*?-->/g, "");
}

export function toHtmlFromText(text) {
  return `<div style="white-space: pre-wrap;">${escapeHtml(text || "")}</div>`;
}

export function displayFormWithFallback(displayFn, htmlBody) {
  if (!htmlBody) {
    displayFn("");
    return;
  }

  try {
    displayFn({ htmlBody });
  } catch {
    displayFn("");
  }
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export function getForwardSubject(subject) {
  const sourceSubject = String(subject || "").trim();
  const prefix = t("commands.notifications.forwardPrefix");

  if (!sourceSubject) {
    return prefix;
  }

  return new RegExp(`^${escapeRegex(prefix)}\\s*`, "i").test(sourceSubject)
    ? sourceSubject
    : `${prefix} ${sourceSubject}`;
}
