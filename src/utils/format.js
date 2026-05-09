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
