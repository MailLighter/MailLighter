/**
 * sendHandler tests verify the core invariant of Phase 1a:
 *   - savings counter must update only when the user actually sends
 *   - the recipientCount used for the multiplier must be the LATEST
 *     count at send time, not the one captured at cleanup
 *   - the send must never be blocked, even on internal failures
 */

const store = new Map();
global.localStorage = {
  getItem: (k) => (store.has(k) ? store.get(k) : null),
  setItem: (k, v) => store.set(k, String(v)),
  removeItem: (k) => store.delete(k),
  clear: () => store.clear(),
};

const { onMessageSend, confirmEagerly } = require("../src/core/lifecycle/sendHandler");
const {
  addPendingEvent,
  _resetPendingForTests,
  getPendingState,
} = require("../src/core/savings/pendingSavings");
const { createStorage, getSavings } = require("../src/core/savings/savingsCalculator");
const { ELEMENT_TYPES } = require("../src/config/constants");

function makePlatform({ composeId = "c1", recipientCount = 1, getRecipientsThrows = false } = {}) {
  return {
    platformName: "outlook",
    async getComposeId() {
      return composeId;
    },
    async getRecipients() {
      if (getRecipientsThrows) throw new Error("getRecipients failed");
      return new Array(recipientCount).fill({ emailAddress: "x@y" });
    },
  };
}

function makeEventArgs() {
  const calls = [];
  return {
    completed: (arg) => calls.push(arg),
    _calls: calls,
  };
}

function pendingEvent(elementType, bytesRemoved, recipientCount) {
  return { elementType, bytesRemoved, recipientCount, timestamp: Date.now(), platform: "outlook" };
}

beforeEach(() => {
  store.clear();
  _resetPendingForTests();
});

describe("sendHandler — pending → send → consume cycle", () => {
  test("confirms pending events into the transmission counter when send completes", async () => {
    addPendingEvent("c1", pendingEvent(ELEMENT_TYPES.IMAGE, 1000, 1));

    const platform = makePlatform({ composeId: "c1", recipientCount: 1 });
    const eventArgs = makeEventArgs();
    const storage = createStorage();

    await onMessageSend(eventArgs, platform, storage);

    const s = getSavings(storage);
    expect(s.transmission.images).toBe(1000);
    expect(getPendingState().count).toBe(0);
  });

  test("uses the LATEST recipient count at send time, not the one captured at cleanup", async () => {
    // Cleanup happened with 3 recipients
    addPendingEvent("c1", pendingEvent(ELEMENT_TYPES.IMAGE, 1000, 3));

    // But by send time the user added 2 recipients (now 5)
    const platform = makePlatform({ composeId: "c1", recipientCount: 5 });
    const eventArgs = makeEventArgs();
    const storage = createStorage();

    await onMessageSend(eventArgs, platform, storage);

    const s = getSavings(storage);
    expect(s.transmission.images).toBe(1000 * 5);
  });

  test("aggregates multiple pending events into one send confirmation", async () => {
    addPendingEvent("c1", pendingEvent(ELEMENT_TYPES.IMAGE, 100, 1));
    addPendingEvent("c1", pendingEvent(ELEMENT_TYPES.REPLY, 200, 1));
    addPendingEvent("c1", pendingEvent(ELEMENT_TYPES.ATTACHMENT, 400, 1));

    const platform = makePlatform({ composeId: "c1", recipientCount: 4 });
    const eventArgs = makeEventArgs();
    const storage = createStorage();

    await onMessageSend(eventArgs, platform, storage);

    const s = getSavings(storage);
    expect(s.transmission.total).toBe((100 + 200 + 400) * 4);
  });

  test("does nothing when no events are pending for the composeId", async () => {
    const platform = makePlatform({ composeId: "c1" });
    const eventArgs = makeEventArgs();
    const storage = createStorage();

    await onMessageSend(eventArgs, platform, storage);

    const s = getSavings(storage);
    expect(s.transmission.total).toBe(0);
  });
});

describe("sendHandler — discard / cleanup-without-send", () => {
  test("a different composeId does not consume another's pending events", async () => {
    addPendingEvent("c1", pendingEvent(ELEMENT_TYPES.IMAGE, 1000, 1));

    // user opens a SECOND compose ("c2"), no pending events there, sends c2.
    const platform = makePlatform({ composeId: "c2", recipientCount: 1 });
    const eventArgs = makeEventArgs();
    const storage = createStorage();

    await onMessageSend(eventArgs, platform, storage);

    // c2 sent: counter unchanged, c1's pending events still queued (eventually
    // expire via TTL since c1 was never sent).
    const s = getSavings(storage);
    expect(s.transmission.total).toBe(0);
    expect(getPendingState().count).toBe(1);
  });
});

describe("confirmEagerly — legacy Mailbox < 1.10 path", () => {
  test("confirms pending events using the captured recipientCount", async () => {
    addPendingEvent("c1", pendingEvent(ELEMENT_TYPES.IMAGE, 1000, 3));
    const storage = createStorage();
    await confirmEagerly("c1", storage);
    expect(getSavings(storage).transmission.images).toBe(1000 * 3);
    expect(getPendingState().count).toBe(0);
  });

  test("confirms multiple pending events in one call", async () => {
    addPendingEvent("c1", pendingEvent(ELEMENT_TYPES.IMAGE, 100, 2));
    addPendingEvent("c1", pendingEvent(ELEMENT_TYPES.REPLY, 200, 2));
    addPendingEvent("c1", pendingEvent(ELEMENT_TYPES.ATTACHMENT, 400, 2));
    const storage = createStorage();
    await confirmEagerly("c1", storage);
    expect(getSavings(storage).transmission.total).toBe((100 + 200 + 400) * 2);
  });

  test("does nothing when no events are pending", async () => {
    const storage = createStorage();
    await confirmEagerly("c1", storage);
    expect(getSavings(storage).transmission.total).toBe(0);
  });

  test("does nothing when composeId is falsy", async () => {
    addPendingEvent("c1", pendingEvent(ELEMENT_TYPES.IMAGE, 500, 1));
    const storage = createStorage();
    await confirmEagerly(null, storage);
    await confirmEagerly("", storage);
    expect(getPendingState().count).toBe(1);
    expect(getSavings(storage).transmission.total).toBe(0);
  });

  test("isolates events between composeIds", async () => {
    addPendingEvent("c1", pendingEvent(ELEMENT_TYPES.IMAGE, 1000, 1));
    addPendingEvent("c2", pendingEvent(ELEMENT_TYPES.IMAGE, 500, 1));
    const storage = createStorage();
    await confirmEagerly("c1", storage);
    expect(getSavings(storage).transmission.images).toBe(1000);
    expect(getPendingState().count).toBe(1);
  });

  test("survives a thrown exception without propagating", async () => {
    const brokenStorage = {
      getItem: () => { throw new Error("storage failure"); },
      setItem: () => { throw new Error("storage failure"); },
    };
    addPendingEvent("c1", pendingEvent(ELEMENT_TYPES.IMAGE, 100, 1));
    await expect(confirmEagerly("c1", brokenStorage)).resolves.toBeUndefined();
  });
});

describe("sendHandler — never blocks the send", () => {
  test("calls eventArgs.completed({ allowEvent: true }) on success", async () => {
    addPendingEvent("c1", pendingEvent(ELEMENT_TYPES.IMAGE, 100, 1));
    const platform = makePlatform({ composeId: "c1" });
    const eventArgs = makeEventArgs();

    await onMessageSend(eventArgs, platform, createStorage());

    expect(eventArgs._calls).toHaveLength(1);
    expect(eventArgs._calls[0]).toEqual({ allowEvent: true });
  });

  test("calls eventArgs.completed({ allowEvent: true }) even when getRecipients throws", async () => {
    addPendingEvent("c1", pendingEvent(ELEMENT_TYPES.IMAGE, 100, 1));
    const platform = makePlatform({ composeId: "c1", getRecipientsThrows: true });
    const eventArgs = makeEventArgs();

    await onMessageSend(eventArgs, platform, createStorage());

    expect(eventArgs._calls).toHaveLength(1);
    expect(eventArgs._calls[0]).toEqual({ allowEvent: true });
  });
});
