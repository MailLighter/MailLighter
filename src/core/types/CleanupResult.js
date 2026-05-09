import { ELEMENT_TYPES } from "../../config/constants";

/**
 * Résultat retourné par un cleaner après exécution.
 * Sert au feedback notification immédiat (count + bytes), indépendamment
 * du compteur Settings qui n'est mis à jour qu'au send.
 */
export function createCleanupResult({ elementType, itemsRemoved, bytesRemoved }) {
  if (!Object.values(ELEMENT_TYPES).includes(elementType)) {
    throw new Error(`Invalid elementType: ${elementType}`);
  }
  if (typeof itemsRemoved !== "number" || itemsRemoved < 0) {
    throw new Error(`Invalid itemsRemoved: ${itemsRemoved}`);
  }
  if (typeof bytesRemoved !== "number" || bytesRemoved < 0) {
    throw new Error(`Invalid bytesRemoved: ${bytesRemoved}`);
  }
  return Object.freeze({
    elementType,
    itemsRemoved,
    bytesRemoved,
  });
}
