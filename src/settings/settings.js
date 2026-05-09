/* global Office, document, window, URLSearchParams, clearTimeout, setTimeout */

import { t } from "../ui/i18n";
import { formatFileSize, escapeHtml } from "../utils/format";
import { MAILLIGHTER_SITE_URL } from "../config/constants";

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

function intParam(params, key) {
  const v = parseInt(params.get(key) || "0", 10);
  return Number.isFinite(v) ? v : 0;
}

Office.onReady(() => {
  const params = new URLSearchParams(window.location.search);
  const ecoEnabled = params.get("ecoMessage") === "1";
  const ecoText = params.get("ecoText") || t("settings.ecoMessageDefault");

  const transImages = intParam(params, "transImages");
  const transReplies = intParam(params, "transReplies");
  const transAttachments = intParam(params, "transAttachments");
  const transTotal = intParam(params, "transTotal");

  const checkbox = document.getElementById("ecoMessageCheckbox");
  const ecoPreview = document.getElementById("ecoPreview");
  const textarea = document.getElementById("ecoMessageTextarea");
  const closeButton = document.getElementById("closeButton");

  // Apply i18n text — page chrome
  document.getElementById("settingsTitle").textContent = t("settings.title");
  document.getElementById("ecoMessageTitle").textContent = t("settings.ecoMessageTitle");
  document.getElementById("ecoMessageDescription").textContent = t(
    "settings.ecoMessageDescription"
  );
  document.getElementById("previewLabel").textContent = t("settings.previewLabel");
  document.getElementById("ecoMessageEditLabel").textContent = t("settings.ecoMessageEditLabel");
  document.getElementById("ecoResetButton").textContent = t("settings.ecoMessageReset");

  // Transmission savings section (raw × recipients)
  document.getElementById("savingsTransmissionTitle").textContent = t(
    "settings.savingsTransmissionTitle"
  );
  document.getElementById("savingsTransmissionHint").textContent = t(
    "settings.savingsTransmissionHint"
  );
  document.getElementById("transSavingsImagesLabel").textContent = t("settings.savingsImages");
  document.getElementById("transSavingsRepliesLabel").textContent = t("settings.savingsReplies");
  document.getElementById("transSavingsAttachmentsLabel").textContent = t(
    "settings.savingsAttachments"
  );
  document.getElementById("transSavingsTotalLabel").textContent = t("settings.savingsTotal");

  closeButton.textContent = t("settings.close");

  // Set initial eco message state
  checkbox.checked = ecoEnabled;
  textarea.value = ecoText;
  updatePreview(ecoText);
  if (ecoEnabled) ecoPreview.classList.add("visible");

  // Display values — transmission savings
  document.getElementById("transSavingsImages").textContent = formatBytes(transImages);
  document.getElementById("transSavingsReplies").textContent = formatBytes(transReplies);
  document.getElementById("transSavingsAttachments").textContent = formatBytes(transAttachments);
  document.getElementById("transSavingsTotal").textContent = formatBytes(transTotal);

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
