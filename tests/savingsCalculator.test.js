const store = new Map();
global.localStorage = {
  getItem: (k) => (store.has(k) ? store.get(k) : null),
  setItem: (k, v) => store.set(k, String(v)),
  removeItem: (k) => store.delete(k),
  clear: () => store.clear(),
};

const {
  createStorage,
  recordConfirmedSavings,
  getSavings,
} = require("../src/core/savings/savingsCalculator");
const { ELEMENT_TYPES } = require("../src/config/constants");

beforeEach(() => store.clear());

function event(elementType, bytesRemoved, recipientCount) {
  return { elementType, bytesRemoved, recipientCount };
}

describe("recordConfirmedSavings", () => {
  test("updates transmission counter", () => {
    const storage = createStorage();
    recordConfirmedSavings(storage, event(ELEMENT_TYPES.IMAGE, 1024, 5));
    const s = getSavings(storage);
    expect(s.transmission.images).toBe(1024 * 5);
  });

  test("ignores zero or negative bytesRemoved", () => {
    const storage = createStorage();
    recordConfirmedSavings(storage, event(ELEMENT_TYPES.IMAGE, 0, 5));
    recordConfirmedSavings(storage, event(ELEMENT_TYPES.IMAGE, -100, 5));
    const s = getSavings(storage);
    expect(s.transmission.images).toBe(0);
  });

  test("treats recipientCount of 0 as 1 (an email being sent has at least one recipient)", () => {
    const storage = createStorage();
    recordConfirmedSavings(storage, event(ELEMENT_TYPES.IMAGE, 100, 0));
    const s = getSavings(storage);
    expect(s.transmission.images).toBe(100);
  });

  test("maps SELECTION elementType onto the replies category", () => {
    const storage = createStorage();
    recordConfirmedSavings(storage, event(ELEMENT_TYPES.SELECTION, 500, 2));
    const s = getSavings(storage);
    expect(s.transmission.replies).toBe(1000);
  });

  test("REPLY and SELECTION accumulate in the same replies category", () => {
    const storage = createStorage();
    recordConfirmedSavings(storage, event(ELEMENT_TYPES.REPLY, 100, 2));
    recordConfirmedSavings(storage, event(ELEMENT_TYPES.SELECTION, 200, 3));
    const s = getSavings(storage);
    expect(s.transmission.replies).toBe(100 * 2 + 200 * 3);
  });

  test("accumulates across categories independently", () => {
    const storage = createStorage();
    recordConfirmedSavings(storage, event(ELEMENT_TYPES.IMAGE, 100, 2));
    recordConfirmedSavings(storage, event(ELEMENT_TYPES.REPLY, 200, 3));
    recordConfirmedSavings(storage, event(ELEMENT_TYPES.ATTACHMENT, 400, 4));
    const s = getSavings(storage);
    expect(s.transmission).toEqual({
      images: 200,
      replies: 600,
      attachments: 1600,
      total: 2400,
    });
  });

  test("ignores unknown elementType silently", () => {
    const storage = createStorage();
    recordConfirmedSavings(storage, { elementType: "bogus", bytesRemoved: 100, recipientCount: 1 });
    const s = getSavings(storage);
    expect(s.transmission.total).toBe(0);
  });
});

describe("getSavings", () => {
  test("returns zeros when nothing has been recorded", () => {
    const s = getSavings(createStorage());
    expect(s.transmission).toEqual({ images: 0, replies: 0, attachments: 0, total: 0 });
  });

  test("recovers from non-numeric stored values", () => {
    store.set("maillighter_savings_transmission_images", "garbage");
    const s = getSavings(createStorage());
    expect(s.transmission.images).toBe(0);
  });
});
