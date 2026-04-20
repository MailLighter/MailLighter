const KEYS = {
  images: "maillighter_savings_images",
  replies: "maillighter_savings_replies",
  attachments: "maillighter_savings_attachments",
};

export function addSavings(category, bytes) {
  if (!bytes || bytes <= 0) return;
  const key = KEYS[category];
  if (!key) return;
  const current = parseInt(localStorage.getItem(key) || "0", 10);
  localStorage.setItem(key, String(current + bytes));
}

export function getSavings() {
  const images = parseInt(localStorage.getItem(KEYS.images) || "0", 10);
  const replies = parseInt(localStorage.getItem(KEYS.replies) || "0", 10);
  const attachments = parseInt(localStorage.getItem(KEYS.attachments) || "0", 10);
  return { images, replies, attachments, total: images + replies + attachments };
}
