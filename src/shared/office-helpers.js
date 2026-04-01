export function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
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
      // Remove javascript: and data: URIs from href and src attributes
      .replace(/(<[^>]*\s)(href|src)\s*=\s*["']?\s*javascript\s*:[^"'>]*/gi, '$1$2=""')
      .replace(/(<[^>]*\s)(href|src)\s*=\s*["']?\s*data\s*:\s*text\/html[^"'>]*/gi, '$1$2=""')
  );
}

export function toHtmlFromText(text) {
  return `<div style="white-space: pre-wrap;">${escapeHtml(text || "")}</div>`;
}
