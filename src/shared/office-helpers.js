export const MAILLIGHTER_SITE_URL = "https://www.maillighter.com";

export function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

export function formatFileSize(bytes, { zeroLabel = "", units } = {}) {
  const labels = units || {
    kilobytes: "KB",
    megabytes: "MB",
    gigabytes: "GB",
    lessThanOne: "< 1 KB",
  };

  if (!bytes || bytes <= 0) {
    return zeroLabel;
  }

  const kilobytes = bytes / 1024;

  if (kilobytes < 1) {
    return labels.lessThanOne;
  }

  if (kilobytes < 1024) {
    return `${Math.round(kilobytes * 100) / 100} ${labels.kilobytes}`;
  }

  if (kilobytes < 1024 * 1024) {
    return `${Math.round((kilobytes / 1024) * 100) / 100} ${labels.megabytes}`;
  }

  return `${Math.round((kilobytes / (1024 * 1024)) * 100) / 100} ${labels.gigabytes}`;
}

export function sanitizeSelectionHtml(html) {
  return (
    html
      // Remove <script> tags and their content
      .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "")
      // Remove HTML comments
      .replace(/<!--[\s\S]*?-->/g, "")
      // Remove dangerous tags entirely: <iframe>, <embed>, <object>, <applet>, <form>
      .replace(/<(iframe|embed|object|applet|form)\b[^>]*>[\s\S]*?<\/\1>/gi, "")
      .replace(/<(iframe|embed|object|applet|form)\b[^>]*\/?>/gi, "")
      // Remove event handler attributes (on*)
      .replace(/(<[^>]*)\s+on\w+\s*=\s*("[^"]*"|'[^']*'|[^\s>]*)/gi, "$1")
      // Remove javascript: URIs from href and src attributes
      .replace(/(<[^>]*\s)(href|src)\s*=\s*["']?\s*javascript\s*:[^"'>]*/gi, '$1$2=""')
      // Remove all data: URIs from href and src attributes (text/html, text/javascript, etc.)
      .replace(/(<[^>]*\s)(href|src)\s*=\s*["']?\s*data\s*:[^"'>]*/gi, '$1$2=""')
  );
}

export function toHtmlFromText(text) {
  return `<div style="white-space: pre-wrap;">${escapeHtml(text || "")}</div>`;
}
