/* global Office, document, window, URLSearchParams */

import { t } from "../shared/i18n";

function formatBytes(bytes) {
  if (!bytes || bytes <= 0) return `0 ${t("units.kilobytes")}`;
  const kb = bytes / 1024;
  if (kb < 1) return t("units.lessThanOne");
  if (kb < 1024) return `${Math.round(kb * 100) / 100} ${t("units.kilobytes")}`;
  if (kb < 1024 * 1024) return `${Math.round((kb / 1024) * 100) / 100} ${t("units.megabytes")}`;
  return `${Math.round((kb / (1024 * 1024)) * 100) / 100} ${t("units.gigabytes")}`;
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function updatePreview(text) {
  const linked = escapeHtml(text).replace(
    /MailLighter/g,
    '<a href="https://www.maillighter.com" target="_blank" style="color:#1b5e20;">MailLighter</a>'
  );
  document.getElementById("ecoPreviewText").innerHTML = linked;
}

Office.onReady(() => {
  const params = new URLSearchParams(window.location.search);
  const ecoEnabled = params.get("ecoMessage") === "1";
  const ecoText = params.get("ecoText") || t("settings.ecoMessageDefault");
  const savImages = parseInt(params.get("savImages") || "0", 10);
  const savReplies = parseInt(params.get("savReplies") || "0", 10);
  const savAttachments = parseInt(params.get("savAttachments") || "0", 10);
  const savTotal = parseInt(params.get("savTotal") || "0", 10);

  const checkbox = document.getElementById("ecoMessageCheckbox");
  const ecoPreview = document.getElementById("ecoPreview");
  const textarea = document.getElementById("ecoMessageTextarea");
  const closeButton = document.getElementById("closeButton");

  // Apply i18n text
  document.getElementById("settingsTitle").textContent = t("settings.title");
  document.getElementById("ecoMessageTitle").textContent = t("settings.ecoMessageTitle");
  document.getElementById("ecoMessageDescription").textContent = t(
    "settings.ecoMessageDescription"
  );
  document.getElementById("previewLabel").textContent = t("settings.previewLabel");
  document.getElementById("ecoMessageEditLabel").textContent = t("settings.ecoMessageEditLabel");
  document.getElementById("ecoResetButton").textContent = t("settings.ecoMessageReset");
  document.getElementById("savingsTitle").textContent = t("settings.savingsTitle");
  document.getElementById("savingsImagesLabel").textContent = t("settings.savingsImages");
  document.getElementById("savingsRepliesLabel").textContent = t("settings.savingsReplies");
  document.getElementById("savingsAttachmentsLabel").textContent = t("settings.savingsAttachments");
  document.getElementById("savingsTotalLabel").textContent = t("settings.savingsTotal");
  closeButton.textContent = t("settings.close");

  // Set initial eco message state
  checkbox.checked = ecoEnabled;
  textarea.value = ecoText;
  updatePreview(ecoText);
  if (ecoEnabled) ecoPreview.classList.add("visible");

  // Display savings values
  document.getElementById("savingsImages").textContent = formatBytes(savImages);
  document.getElementById("savingsReplies").textContent = formatBytes(savReplies);
  document.getElementById("savingsAttachments").textContent = formatBytes(savAttachments);
  document.getElementById("savingsTotal").textContent = formatBytes(savTotal);

  checkbox.addEventListener("change", () => {
    if (checkbox.checked) {
      ecoPreview.classList.add("visible");
    } else {
      ecoPreview.classList.remove("visible");
    }
    Office.context.ui.messageParent(JSON.stringify({ ecoMessageEnabled: checkbox.checked }));
  });

  textarea.addEventListener("input", () => {
    updatePreview(textarea.value);
    Office.context.ui.messageParent(JSON.stringify({ ecoMessageText: textarea.value }));
  });

  document.getElementById("ecoResetButton").addEventListener("click", () => {
    const defaultText = t("settings.ecoMessageDefault");
    textarea.value = defaultText;
    updatePreview(defaultText);
    Office.context.ui.messageParent(JSON.stringify({ ecoMessageText: defaultText }));
  });

  closeButton.addEventListener("click", () => {
    // Ask the parent to close us (works on Outlook Web).  Desktop cleans up
    // the dialog handle after event.completed() so dialog.close() from the
    // parent throws; fall back to window.close() which works on WebView2.
    try {
      Office.context.ui.messageParent(JSON.stringify({ action: "close" }));
    } catch {
      // ignore
    }
    window.close();
  });
});
