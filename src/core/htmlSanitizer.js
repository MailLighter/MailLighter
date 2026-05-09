import { escapeHtml } from "../utils/format";

/**
 * Strips potentially dangerous HTML constructs from selection content
 * before re-injection into an email body.
 */
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
      // Remove all data: URIs from href and src attributes
      .replace(/(<[^>]*\s)(href|src)\s*=\s*["']?\s*data\s*:[^"'>]*/gi, '$1$2=""')
  );
}

export function toHtmlFromText(text) {
  return `<div style="white-space: pre-wrap;">${escapeHtml(text || "")}</div>`;
}
