/* global Office, document, window, URLSearchParams, clearTimeout, setTimeout */

import { t } from "../shared/i18n";
import { escapeHtml, formatFileSize, MAILLIGHTER_SITE_URL } from "../shared/office-helpers";

function unitLabels() {
  return {
    kilobytes: t("units.kilobytes"),
    megabytes: t("units.megabytes"),
    gigabytes: t("units.gigabytes"),
    lessThanOne: t("units.lessThanOne"),
  };
}

function formatBytes(bytes) {
  return formatFileSize(bytes, {
    zeroLabel: `0 ${t("units.kilobytes")}`,
    units: unitLabels(),
  });
}

function updatePreview(text) {
  const linked = escapeHtml(text).replace(
    /MailLighter/g,
    `<a href="${MAILLIGHTER_SITE_URL}" target="_blank" style="color:#1b5e20;">MailLighter</a>`
  );
  // NOTE: `linked` is safe here because it is produced by escapeHtml() followed
  // by a single .replace() whose href is a hardcoded constant (MAILLIGHTER_SITE_URL).
  // Any future modification that introduces user-controlled content into `linked`
  // before this assignment must go through escapeHtml() first.
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

  let ecoTextDebounceTimer = null;
  textarea.addEventListener("input", () => {
    updatePreview(textarea.value);
    clearTimeout(ecoTextDebounceTimer);
    ecoTextDebounceTimer = setTimeout(() => {
      Office.context.ui.messageParent(JSON.stringify({ ecoMessageText: textarea.value }));
    }, 300);
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
