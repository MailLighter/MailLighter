// Phase 1a: aucun flag actif. Les flags Enterprise (ENTERPRISE_MODE,
// TELEMETRY_LOCAL_BUFFER, TELEMETRY_NETWORK, SHOW_ENTERPRISE_BANNER) seront
// ajoutés en Phase 1b, ainsi que le helper fromEnv() pour les lire depuis
// process.env.
export const FEATURE_FLAGS = Object.freeze({});

export function validateFeatureFlags() {
  // Phase 1a: noop. Sera enrichi en Phase 1b avec les contraintes inter-flags
  // (par ex. TELEMETRY_NETWORK requires ENTERPRISE_MODE).
}
