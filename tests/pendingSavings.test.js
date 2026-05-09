const {
  addPendingEvent,
  consumePendingEvents,
  discardPendingEvents,
  getPendingState,
  purgeStale,
  _resetPendingForTests,
} = require("../src/core/savings/pendingSavings");

beforeEach(() => _resetPendingForTests());

const sampleEvent = (bytes = 100) => ({
  elementType: "image",
  bytesRemoved: bytes,
  recipientCount: 1,
  timestamp: Date.now(),
  platform: "outlook",
});

describe("addPendingEvent + consumePendingEvents", () => {
  test("queues events under a composeId and consume returns them in order", () => {
    addPendingEvent("c1", sampleEvent(100));
    addPendingEvent("c1", sampleEvent(200));
    const events = consumePendingEvents("c1");
    expect(events).toHaveLength(2);
    expect(events[0].bytesRemoved).toBe(100);
    expect(events[1].bytesRemoved).toBe(200);
  });

  test("consume clears the queue for that composeId", () => {
    addPendingEvent("c1", sampleEvent(100));
    consumePendingEvents("c1");
    expect(consumePendingEvents("c1")).toEqual([]);
  });

  test("isolates events between distinct composeIds", () => {
    addPendingEvent("c1", sampleEvent(100));
    addPendingEvent("c2", sampleEvent(200));
    expect(consumePendingEvents("c1")[0].bytesRemoved).toBe(100);
    expect(consumePendingEvents("c2")[0].bytesRemoved).toBe(200);
  });

  test("ignores events without a composeId", () => {
    addPendingEvent("", sampleEvent(100));
    addPendingEvent(null, sampleEvent(200));
    addPendingEvent(undefined, sampleEvent(300));
    expect(getPendingState().count).toBe(0);
  });
});

describe("discardPendingEvents", () => {
  test("clears the queue without returning events", () => {
    addPendingEvent("c1", sampleEvent(100));
    addPendingEvent("c1", sampleEvent(200));
    discardPendingEvents("c1");
    expect(consumePendingEvents("c1")).toEqual([]);
  });
});

describe("purgeStale", () => {
  test("removes entries older than the supplied TTL", () => {
    addPendingEvent("old", sampleEvent());
    // A negative TTL means "everything is past the limit", so all entries
    // are purged. This avoids needing fake timers to age the createdAt.
    purgeStale(-1);
    expect(getPendingState().count).toBe(0);
  });

  test("keeps entries within the TTL window", () => {
    addPendingEvent("recent", sampleEvent());
    purgeStale(60_000);
    expect(getPendingState().count).toBe(1);
  });
});
