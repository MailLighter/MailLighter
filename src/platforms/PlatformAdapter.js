/* eslint-disable @typescript-eslint/no-unused-vars */

/**
 * Interface PlatformAdapter — contrat que chaque plateforme (Outlook, Gmail)
 * doit implémenter. Les modules `core/` ne doivent JAMAIS dépendre de la
 * plateforme directement, uniquement de cette interface.
 *
 * getUserIdHash N'EXISTE PAS et ne sera jamais ajouté (invariant RGPD).
 *
 * Les paramètres des stubs ci-dessous documentent la signature attendue
 * des implémentations concrètes. Ils sont ignorés ici (eslint-disable).
 */
export class PlatformAdapter {
  async getCurrentItem() {
    throw new Error("getCurrentItem must be implemented");
  }
  async getRecipients() {
    throw new Error("getRecipients must be implemented");
  }
  async getBodyHtml() {
    throw new Error("getBodyHtml must be implemented");
  }
  async setBodyHtml(html) {
    throw new Error("setBodyHtml must be implemented");
  }
  async prependBodyHtml(html) {
    throw new Error("prependBodyHtml must be implemented");
  }
  async getSelectedHtml() {
    throw new Error("getSelectedHtml must be implemented");
  }
  async listAttachments() {
    throw new Error("listAttachments must be implemented");
  }
  async removeAttachment(attachmentId) {
    throw new Error("removeAttachment must be implemented");
  }
  getLocale() {
    throw new Error("getLocale must be implemented");
  }
  notify(message, kind) {
    throw new Error("notify must be implemented");
  }
}
