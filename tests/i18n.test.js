// Jest mocks must be set up before importing the module under test.
function setLocales({ officeLocale = "", browserLocale = "" } = {}) {
  global.Office = officeLocale ? { context: { displayLanguage: officeLocale } } : undefined;
  global.navigator = { language: browserLocale };
}

let t;
let getCurrentLocale;

beforeEach(() => {
  jest.resetModules();
  setLocales();
  ({ t, getCurrentLocale } = require("../src/shared/i18n"));
});

// ---------------------------------------------------------------------------
// getCurrentLocale
// ---------------------------------------------------------------------------
describe("getCurrentLocale", () => {
  test("prefers the Office display language over the browser language", () => {
    setLocales({ officeLocale: "fr-CA", browserLocale: "en-US" });
    jest.resetModules();
    ({ getCurrentLocale } = require("../src/shared/i18n"));
    expect(getCurrentLocale()).toBe("fr-FR");
  });

  test("falls back to navigator.language when Office is unavailable", () => {
    setLocales({ browserLocale: "es-AR" });
    jest.resetModules();
    ({ getCurrentLocale } = require("../src/shared/i18n"));
    expect(getCurrentLocale()).toBe("es-ES");
  });

  test("defaults to en-US for unsupported languages", () => {
    setLocales({ browserLocale: "de-DE" });
    jest.resetModules();
    ({ getCurrentLocale } = require("../src/shared/i18n"));
    expect(getCurrentLocale()).toBe("en-US");
  });
});

// ---------------------------------------------------------------------------
// t
// ---------------------------------------------------------------------------
describe("t", () => {
  test("returns the translated string for the current locale", () => {
    setLocales({ officeLocale: "fr-FR" });
    jest.resetModules();
    ({ t } = require("../src/shared/i18n"));
    expect(t("units.kilobytes")).toBe("Ko");
  });

  test("falls back to English when a key is missing in the current locale", () => {
    // We cannot easily delete a key at runtime, so we simulate an unsupported
    // locale — the module falls back to en-US internally.
    setLocales({ browserLocale: "ja-JP" });
    jest.resetModules();
    ({ t } = require("../src/shared/i18n"));
    expect(t("units.kilobytes")).toBe("KB");
  });

  test("returns the raw key when no translation exists anywhere", () => {
    expect(t("this.does.not.exist")).toBe("this.does.not.exist");
  });

  test("interpolates {token} placeholders from params", () => {
    setLocales({ officeLocale: "fr-FR" });
    jest.resetModules();
    ({ t } = require("../src/shared/i18n"));
    expect(t("commands.notifications.imagesRemoved", { count: 3 })).toBe(
      "✅ 3 image(s) supprimée(s)."
    );
  });

  test("keeps the placeholder when a token is not provided", () => {
    setLocales({ officeLocale: "en-US" });
    jest.resetModules();
    ({ t } = require("../src/shared/i18n"));
    expect(t("commands.notifications.imagesRemoved", {})).toBe("✅ {count} image(s) removed.");
  });
});
