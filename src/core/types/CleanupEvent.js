import { ELEMENT_TYPES } from "../../config/constants";

/**
 * Construit l'objet passé à recordConfirmedSavings après un cleanup.
 * Le recipientCount est capturé au moment du clic — c'est lui qui multiplie
 * bytesRemoved pour produire l'économie de transmission.
 */
export function createCleanupEvent({ elementType, bytesRemoved, recipientCount }) {
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
  });
}
