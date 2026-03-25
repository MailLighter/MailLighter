export function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

export function sanitizeSelectionHtml(html) {
  return html
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "")
    .replace(/<!--[\s\S]*?-->/g, "");
}

export function toHtmlFromText(text) {
  return `<div style="white-space: pre-wrap;">${escapeHtml(text || "")}</div>`;
}
