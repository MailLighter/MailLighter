/* global Office */

import { t } from "../shared/i18n";

Office.onReady(() => {
  const params = new URLSearchParams(window.location.search);
  const ecoEnabled = params.get("ecoMessage") === "1";

  const checkbox = document.getElementById("ecoMessageCheckbox");
  const ecoPreview = document.getElementById("ecoPreview");
  const closeButton = document.getElementById("closeButton");

  // Apply i18n text
  document.getElementById("settingsTitle").textContent = `⚙️ ${t("settings.title")}`;
  document.getElementById("ecoMessageTitle").textContent = t("settings.ecoMessageTitle");
  document.getElementById("ecoMessageDescription").textContent = t("settings.ecoMessageDescription");
  document.getElementById("previewLabel").textContent = t("settings.previewLabel");
  closeButton.textContent = t("settings.close");

  // Set initial state
  checkbox.checked = ecoEnabled;
  if (ecoEnabled) ecoPreview.classList.add("visible");

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
