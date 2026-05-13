export const APP_VERSION = "1.0.0";

export const MAILLIGHTER_SITE_URL = "https://www.maillighter.com";

export const CLEANUP_DEFAULTS = Object.freeze({
  KEEP_REPLIES_COUNT: 2,
  MIN_BYTES_TO_REPORT: 100,
});

export const STORAGE_KEYS = Object.freeze({
  USER_SAVINGS_TRANSMISSION_IMAGES: "maillighter_savings_transmission_images",
  USER_SAVINGS_TRANSMISSION_REPLIES: "maillighter_savings_transmission_replies",
  USER_SAVINGS_TRANSMISSION_ATTACHMENTS: "maillighter_savings_transmission_attachments",
  ECO_MESSAGE_ENABLED: "maillighter_eco_message",
  ECO_MESSAGE_TEXT: "maillighter_eco_message_text",
});

export const ELEMENT_TYPES = Object.freeze({
  IMAGE: "image",
  ATTACHMENT: "attachment",
  REPLY: "reply",
  SELECTION: "selection",
});

// Mapping ELEMENT_TYPES → catégorie de stockage Settings.
// "selection" est rangé sous "replies" car les deux concernent du texte/HTML
// supprimé (cohérent avec l'affichage existant "Text & replies").
export const ELEMENT_TYPE_TO_CATEGORY = Object.freeze({
  [ELEMENT_TYPES.IMAGE]: "images",
  [ELEMENT_TYPES.ATTACHMENT]: "attachments",
  [ELEMENT_TYPES.REPLY]: "replies",
  [ELEMENT_TYPES.SELECTION]: "replies",
});

export const PLATFORMS = Object.freeze({
  OUTLOOK: "outlook",
  GMAIL: "gmail",
});
