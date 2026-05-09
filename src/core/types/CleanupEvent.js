import { ELEMENT_TYPES, PLATFORMS } from "../../config/constants";

/**
 * Représente une action de nettoyage candidate, en attente de confirmation
 * par l'envoi effectif de l'email.
 *
 * Le recipientCount est capturé au cleanup (valeur indicative) puis
 * RECALCULÉ au moment du send par sendHandler avant d'être utilisé pour
 * mettre à jour le compteur Settings (bytesRemoved × recipientCount).
 */
export function createCleanupEvent({ elementType, bytesRemoved, recipientCount, platform }) {
  if (!Object.values(ELEMENT_TYPES).includes(elementType)) {
    throw new Error(`Invalid elementType: ${elementType}`);
  }
  if (typeof bytesRemoved !== "number" || bytesRemoved < 0) {
    throw new Error(`Invalid bytesRemoved: ${bytesRemoved}`);
  }
  if (typeof recipientCount !== "number" || recipientCount < 0) {
    throw new Error(`Invalid recipientCount: ${recipientCount}`);
  }
  return Object.freeze({
    elementType,
    bytesRemoved,
    recipientCount,
    timestamp: Date.now(),
    platform: platform || PLATFORMS.OUTLOOK,
  });
}
