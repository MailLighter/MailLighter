/* global Office, navigator */

import enUsTranslations from "../i18n/en-US.json";
import frFrTranslations from "../i18n/fr-FR.json";
import esEsTranslations from "../i18n/es-ES.json";

const TRANSLATIONS = {
  "en-US": enUsTranslations,
  "fr-FR": frFrTranslations,
  "es-ES": esEsTranslations,
};

function getNestedValue(source, path) {
  return path
    .split(".")
    .reduce((acc, key) => (acc && acc[key] !== undefined ? acc[key] : undefined), source);
}

export function getCurrentLocale() {
  const officeLocale =
    typeof Office !== "undefined" && Office.context ? Office.context.displayLanguage : "";
  const browserLocale = typeof navigator !== "undefined" ? navigator.language : "";
  const locale = (officeLocale || browserLocale || "en-US").toLowerCase();

  if (locale.startsWith("fr")) {
    return "fr-FR";
  }

  if (locale.startsWith("es")) {
    return "es-ES";
  }

  return "en-US";
}

export function t(key, params = {}) {
  const locale = getCurrentLocale();
  const localizedValue = getNestedValue(TRANSLATIONS[locale], key);
  const fallbackValue = getNestedValue(TRANSLATIONS["en-US"], key);
  const template = localizedValue || fallbackValue || key;

  return template.replace(/\{(\w+)\}/g, (_, token) => {
    return params[token] !== undefined ? String(params[token]) : `{${token}}`;
  });
}
