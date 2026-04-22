const store = new Map();
global.localStorage = {
  getItem: (k) => (store.has(k) ? store.get(k) : null),
  setItem: (k, v) => store.set(k, String(v)),
  removeItem: (k) => store.delete(k),
  clear: () => store.clear(),
};

const { addSavings, getSavings } = require("../src/shared/savings-storage");

beforeEach(() => store.clear());

// ---------------------------------------------------------------------------
// addSavings
// ---------------------------------------------------------------------------
describe("addSavings", () => {
  test("accumulates bytes for a category", () => {
    addSavings("images", 1024);
    addSavings("images", 2048);
    expect(getSavings().images).toBe(3072);
  });

  test("ignores non-positive values", () => {
    addSavings("images", 0);
    addSavings("images", -500);
    addSavings("images", null);
    addSavings("images", undefined);
    expect(getSavings().images).toBe(0);
  });

  test("ignores unknown categories", () => {
    addSavings("bogus", 1024);
    const s = getSavings();
    expect(s.images).toBe(0);
    expect(s.replies).toBe(0);
    expect(s.attachments).toBe(0);
    expect(s.total).toBe(0);
  });

  test("supports the three known categories independently", () => {
    addSavings("images", 100);
    addSavings("replies", 200);
    addSavings("attachments", 300);
    const s = getSavings();
    expect(s).toEqual({ images: 100, replies: 200, attachments: 300, total: 600 });
  });
});

// ---------------------------------------------------------------------------
// getSavings
// ---------------------------------------------------------------------------
describe("getSavings", () => {
  test("returns zeros when nothing is stored", () => {
    expect(getSavings()).toEqual({ images: 0, replies: 0, attachments: 0, total: 0 });
  });

  test("total equals the sum of the three categories", () => {
    addSavings("images", 1000);
    addSavings("replies", 2000);
    addSavings("attachments", 4000);
    const s = getSavings();
    expect(s.total).toBe(s.images + s.replies + s.attachments);
  });

  test("recovers gracefully from non-numeric storage values", () => {
    store.set("maillighter_savings_images", "not-a-number");
    expect(getSavings().images).toBe(0);
  });
});
