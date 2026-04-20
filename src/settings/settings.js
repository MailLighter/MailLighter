/* global Office */

import { t } from "../shared/i18n";

function formatBytes(bytes) {
  if (!bytes || bytes <= 0) return `0 ${t("units.kilobytes")}`;
  const kb = bytes / 1024;
  if (kb < 1) return t("units.lessThanOne");
  if (kb < 1024) return `${Math.round(kb * 100) / 100} ${t("units.kilobytes")}`;
  if (kb < 1024 * 1024) return `${Math.round((kb / 1024) * 100) / 100} ${t("units.megabytes")}`;
  return `${Math.round((kb / (1024 * 1024)) * 100) / 100} ${t("units.gigabytes")}`;
}

Office.onReady(() => {
  const params = new URLSearchParams(window.location.search);
  const ecoEnabled = params.get("ecoMessage") === "1";
  const savImages = parseInt(params.get("savImages") || "0", 10);
  const savReplies = parseInt(params.get("savReplies") || "0", 10);
  const savAttachments = parseInt(params.get("savAttachments") || "0", 10);
  const savTotal = parseInt(params.get("savTotal") || "0", 10);

  const checkbox = document.getElementById("ecoMessageCheckbox");
  const ecoPreview = document.getElementById("ecoPreview");
  const closeButton = document.getElementById("closeButton");

  // Apply i18n text
  document.getElementById("settingsTitle").textContent = `⚙️ ${t("settings.title")}`;
  document.getElementById("ecoMessageTitle").textContent = t("settings.ecoMessageTitle");
  document.getElementById("ecoMessageDescription").textContent = t("settings.ecoMessageDescription");
  document.getElementById("previewLabel").textContent = t("settings.previewLabel");
  document.getElementById("savingsTitle").textContent = `📊 ${t("settings.savingsTitle")}`;
  document.getElementById("savingsImagesLabel").textContent = t("settings.savingsImages");
  document.getElementById("savingsRepliesLabel").textContent = t("settings.savingsReplies");
  document.getElementById("savingsAttachmentsLabel").textContent = t("settings.savingsAttachments");
  document.getElementById("savingsTotalLabel").textContent = t("settings.savingsTotal");
  closeButton.textContent = t("settings.close");

  // Set initial eco message state
  checkbox.checked = ecoEnabled;
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

  closeButton.addEventListener("click", () => {
    Office.context.ui.messageParent(JSON.stringify({ action: "close" }));
  });
});
