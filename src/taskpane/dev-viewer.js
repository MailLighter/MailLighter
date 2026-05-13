/* global Office */
import { findReplySeparators } from "../core/replyDetection";

let lastHtml = "";

Office.onReady(() => {
  document.getElementById("btnRefresh").addEventListener("click", refresh);
  document.getElementById("btnCopy").addEventListener("click", copyToClipboard);
  document.getElementById("btnCopySep").addEventListener("click", copySeparators);
});

function refresh() {
  const body =
    Office.context.mailbox.item && Office.context.mailbox.item.body;
  if (!body || typeof body.getAsync !== "function") {
    setStatus("body.getAsync not available in this context.");
    return;
  }
  body.getAsync(Office.CoercionType.Html, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("Error: " + (result.error ? result.error.message : "unknown"));
      return;
    }
    lastHtml = result.value || "";
    document.getElementById("htmlOutput").textContent = lastHtml;
    setStatus(
      "Loaded " + lastHtml.length + " chars — " + new Date().toLocaleTimeString()
    );
    showSeparators(lastHtml);
  });
}

function showSeparators(html) {
  const sepList = document.getElementById("sepList");
  const positions = findReplySeparators(html);
  if (positions.length === 0) {
    sepList.innerHTML = '<span class="sep-none">No separators detected.</span>';
    return;
  }
  sepList.innerHTML = positions
    .map((pos, i) => {
      const before = html.substring(Math.max(0, pos - 20), pos);
      const after = html.substring(pos, Math.min(html.length, pos + 120));
      const excerpt =
        escapeHtml(before) +
        '<b style="color:#d83b01">|CUT|</b>' +
        escapeHtml(after);
      return (
        '<div class="sep-item">' +
        '<div class="label">Separator ' +
        (i + 1) +
        " — position " +
        pos +
        "</div>" +
        '<div class="excerpt">' +
        excerpt +
        "</div>" +
        "</div>"
      );
    })
    .join("");
}

function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function copyToClipboard() {
  if (!lastHtml) {
    setStatus("Nothing to copy — click Refresh first.");
    return;
  }
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard
      .writeText(lastHtml)
      .then(() => setStatus("Copied to clipboard!"))
      .catch(() => fallbackCopy(lastHtml, "Copied to clipboard!"));
  } else {
    fallbackCopy(lastHtml, "Copied to clipboard!");
  }
}

function copySeparators() {
  const text = document.getElementById("sepList").innerText || "";
  if (!text || text === "No data yet." || text === "No separators detected.") {
    setStatus("No separator data to copy.");
    return;
  }
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard
      .writeText(text)
      .then(() => setStatus("Separators copied!"))
      .catch(() => fallbackCopy(text, "Separators copied!"));
  } else {
    fallbackCopy(text, "Separators copied!");
  }
}

function fallbackCopy(text, successMsg) {
  const ta = document.createElement("textarea");
  ta.value = text;
  ta.style.cssText = "position:fixed;opacity:0";
  document.body.appendChild(ta);
  ta.select();
  try {
    const ok = document.execCommand("copy");
    setStatus(ok ? successMsg + " (fallback)" : "Copy failed — select the HTML manually.");
  } catch {
    setStatus("Copy failed — select the HTML manually.");
  }
  document.body.removeChild(ta);
}

function setStatus(msg) {
  document.getElementById("status").textContent = msg;
}
