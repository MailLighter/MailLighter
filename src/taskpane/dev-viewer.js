/* global Office, document, navigator */

import { findReplySeparators } from "../shared/reply-detection";

Office.onReady(() => {
  document.getElementById("btnRefresh").addEventListener("click", refresh);
  document.getElementById("btnCopy").addEventListener("click", copyToClipboard);
  document.getElementById("btnCopySep").addEventListener("click", copySeparators);
});

let lastHtml = "";

function setStatus(msg) {
  document.getElementById("status").textContent = msg;
}

function refresh() {
  const body = Office.context.mailbox.item && Office.context.mailbox.item.body;
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
    const output = document.getElementById("htmlOutput");
    output.textContent = lastHtml;
    setStatus("Loaded " + lastHtml.length + " chars — " + new Date().toLocaleTimeString());
    showSeparators(lastHtml);
  });
}

function showSeparators(html) {
  const container = document.getElementById("sepList");
  const positions = findReplySeparators(html);

  if (positions.length === 0) {
    container.innerHTML = '<span class="sep-none">No separators detected.</span>';
    return;
  }

  const items = positions.map((pos, i) => {
    const start = Math.max(0, pos - 20);
    const end = Math.min(html.length, pos + 120);
    const before = html.substring(start, pos);
    const after = html.substring(pos, end);
    const excerpt = escapeHtml(before) + '<b style="color:#d83b01">|CUT|</b>' + escapeHtml(after);
    return (
      '<div class="sep-item">' +
      '<div class="label">Separator ' +
      (i + 1) +
      " — position " +
      pos +
      "</div>" +
      excerpt +
      "</div>"
    );
  });

  container.innerHTML = items.join("");
}

function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function copySeparators() {
  const text = document.getElementById("sepList").innerText;
  if (!text || text === "No data yet." || text === "No separators detected.") {
    setStatus("No separator data to copy.");
    return;
  }
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard
      .writeText(text)
      .then(() => setStatus("Separators copied!"))
      .catch(() => {
        const ta = document.createElement("textarea");
        ta.value = text;
        ta.style.position = "fixed";
        ta.style.opacity = "0";
        document.body.appendChild(ta);
        ta.select();
        try {
          document.execCommand("copy");
          setStatus("Separators copied! (fallback)");
        } catch {
          setStatus("Copy failed.");
        }
        document.body.removeChild(ta);
      });
  }
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
      .catch(() => fallbackCopy());
  } else {
    fallbackCopy();
  }
}

function fallbackCopy() {
  const textarea = document.createElement("textarea");
  textarea.value = lastHtml;
  textarea.style.position = "fixed";
  textarea.style.opacity = "0";
  document.body.appendChild(textarea);
  textarea.select();
  try {
    document.execCommand("copy");
    setStatus("Copied to clipboard! (fallback)");
  } catch {
    setStatus("Copy failed — select the HTML manually.");
  }
  document.body.removeChild(textarea);
}
