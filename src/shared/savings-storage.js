/* global localStorage */

const KEYS = {
  images: "maillighter_savings_images",
  replies: "maillighter_savings_replies",
  attachments: "maillighter_savings_attachments",
};

function readCount(key) {
  const parsed = parseInt(localStorage.getItem(key) || "0", 10);
  return Number.isFinite(parsed) ? parsed : 0;
}

export function addSavings(category, bytes) {
  if (!bytes || bytes <= 0) return;
  const key = KEYS[category];
  if (!key) return;
  localStorage.setItem(key, String(readCount(key) + bytes));
}

export function getSavings() {
  const images = readCount(KEYS.images);
  const replies = readCount(KEYS.replies);
  const attachments = readCount(KEYS.attachments);
  return { images, replies, attachments, total: images + replies + attachments };
}
